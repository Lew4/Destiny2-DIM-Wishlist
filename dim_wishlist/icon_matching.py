"""Normalize embedded perk icons and match them against Bungie artwork."""

from __future__ import annotations

import hashlib
import urllib.error
import urllib.request
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
from PIL import Image

from .icon_config import IconBuilderConfig
from .icon_models import GlobalIconResolution, IconContext, ImageSignature, OfficialVisual
from .manifest import ManifestIndex
from .models import InventoryItem
from .utils import clean_text, norm_name
from .icon_xlsx import sha256_bytes


def _resample_lanczos() -> Any:
    return getattr(Image, "Resampling", Image).LANCZOS


def _content_bbox(image: Image.Image) -> Tuple[int, int, int, int]:
    rgba = np.asarray(image.convert("RGBA"), dtype=np.uint8)
    alpha = rgba[..., 3]
    rgb = rgba[..., :3]
    luminance = 0.2126 * rgb[..., 0] + 0.7152 * rgb[..., 1] + 0.0722 * rgb[..., 2]
    mask = (alpha > 2) & ((luminance > 2) | (alpha > 20))
    ys, xs = np.where(mask)
    if len(xs) == 0:
        return 0, 0, image.width, image.height
    return int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1


def _canonical_rgba(data: bytes, config: IconBuilderConfig) -> np.ndarray:
    image = Image.open(BytesIO(data)).convert("RGBA")
    image = image.crop(_content_bbox(image))
    side = max(image.width, image.height) + 2 * config.content_padding
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.alpha_composite(image, ((side - image.width) // 2, (side - image.height) // 2))
    canvas = canvas.resize(
        (config.normalized_icon_size, config.normalized_icon_size), _resample_lanczos()
    )
    rgba = np.asarray(canvas, dtype=np.uint8).copy()
    alpha = rgba[..., 3:4].astype(np.float32) / 255.0
    rgba[..., :3] = np.round(rgba[..., :3].astype(np.float32) * alpha).astype(np.uint8)
    rgba[rgba[..., 3] == 0, :3] = 0
    return rgba


def _edge_map(values: np.ndarray) -> np.ndarray:
    gradient_y, gradient_x = np.gradient(values.astype(np.float32))
    edge = np.sqrt(gradient_x * gradient_x + gradient_y * gradient_y)
    maximum = float(edge.max())
    if maximum > 1e-9:
        edge /= maximum
    return edge.astype(np.float32)


def _dhash(values: np.ndarray) -> int:
    image = Image.fromarray(np.clip(values * 255.0, 0, 255).astype(np.uint8), mode="L")
    pixels = np.asarray(image.resize((9, 8), _resample_lanczos()), dtype=np.uint8)
    result = 0
    for index, value in enumerate((pixels[:, :-1] > pixels[:, 1:]).ravel()):
        if bool(value):
            result |= 1 << index
    return result


def signature_from_bytes(data: bytes, config: IconBuilderConfig) -> ImageSignature:
    rgba = _canonical_rgba(data, config)
    alpha = rgba[..., 3].astype(np.float32) / 255.0
    rgb = rgba[..., :3].astype(np.float32) / 255.0
    gray = (
        0.2126 * rgb[..., 0] + 0.7152 * rgb[..., 1] + 0.0722 * rgb[..., 2]
    ).astype(np.float32)
    shape = np.maximum(alpha, gray)
    edge = _edge_map(shape)
    fast = Image.fromarray(np.clip(shape * 255.0, 0, 255).astype(np.uint8), mode="L")
    return ImageSignature(
        file_sha256=sha256_bytes(data),
        canonical_sha256=hashlib.sha256(rgba.tobytes()).hexdigest(),
        rgba=rgba,
        alpha=alpha,
        gray=gray,
        edge=edge,
        fast_values=np.asarray(fast.resize((24, 24), _resample_lanczos()), dtype=np.float32) / 255.0,
        dhash=_dhash(shape),
    )


def _shift_zero(array: np.ndarray, dx: int, dy: int) -> np.ndarray:
    output = np.zeros_like(array)
    height, width = array.shape
    source_x0 = max(0, -dx)
    source_x1 = min(width, width - dx) if dx >= 0 else width
    target_x0 = max(0, dx)
    target_x1 = target_x0 + source_x1 - source_x0
    source_y0 = max(0, -dy)
    source_y1 = min(height, height - dy) if dy >= 0 else height
    target_y0 = max(0, dy)
    target_y1 = target_y0 + source_y1 - source_y0
    if source_x1 > source_x0 and source_y1 > source_y0:
        output[target_y0:target_y1, target_x0:target_x1] = array[
            source_y0:source_y1, source_x0:source_x1
        ]
    return output


def _binary_iou(a: np.ndarray, b: np.ndarray, threshold: float = 0.08) -> float:
    aa, bb = a > threshold, b > threshold
    union = int(np.logical_or(aa, bb).sum())
    return 1.0 if union == 0 else float(np.logical_and(aa, bb).sum()) / union


def _cosine_similarity(a: np.ndarray, b: np.ndarray) -> float:
    av, bv = a.ravel().astype(np.float64), b.ravel().astype(np.float64)
    denominator = float(np.linalg.norm(av) * np.linalg.norm(bv))
    if denominator <= 1e-12:
        return 1.0 if np.allclose(av, bv) else 0.0
    return max(0.0, min(1.0, float(np.dot(av, bv) / denominator)))


def _block_ssim(a: np.ndarray, b: np.ndarray, block: int = 16) -> float:
    c1, c2 = 0.01**2, 0.03**2
    height = (a.shape[0] // block) * block
    width = (a.shape[1] // block) * block
    if height == 0 or width == 0:
        return 0.0
    aa = a[:height, :width].astype(np.float64).reshape(
        height // block, block, width // block, block
    )
    bb = b[:height, :width].astype(np.float64).reshape(
        height // block, block, width // block, block
    )
    mean_a, mean_b = aa.mean(axis=(1, 3)), bb.mean(axis=(1, 3))
    variance_a, variance_b = aa.var(axis=(1, 3)), bb.var(axis=(1, 3))
    covariance = (
        (aa - mean_a[:, None, :, None]) * (bb - mean_b[:, None, :, None])
    ).mean(axis=(1, 3))
    numerator = (2 * mean_a * mean_b + c1) * (2 * covariance + c2)
    denominator = (mean_a * mean_a + mean_b * mean_b + c1) * (
        variance_a + variance_b + c2
    )
    values = np.where(denominator > 1e-12, numerator / denominator, 1.0)
    return max(0.0, min(1.0, float(values.mean())))


def fast_image_similarity(a: ImageSignature, b: ImageSignature) -> float:
    hash_score = 1.0 - bin(a.dhash ^ b.dhash).count("1") / 64.0
    return 0.82 * _cosine_similarity(a.fast_values, b.fast_values) + 0.18 * hash_score


def image_similarity(
    a: ImageSignature,
    b: ImageSignature,
    config: IconBuilderConfig,
) -> float:
    if a.file_sha256 == b.file_sha256 or a.canonical_sha256 == b.canonical_sha256:
        return 1.0
    hash_score = 1.0 - bin(a.dhash ^ b.dhash).count("1") / 64.0
    best_shift, best_alignment = (0, 0), -1.0
    translation = config.max_translation_pixels
    for dy in range(-translation, translation + 1):
        for dx in range(-translation, translation + 1):
            shifted_alpha = _shift_zero(b.alpha, dx, dy)
            alignment = (
                0.62 * _cosine_similarity(a.alpha, shifted_alpha)
                + 0.38 * _binary_iou(a.alpha, shifted_alpha)
            )
            if alignment > best_alignment:
                best_alignment, best_shift = alignment, (dx, dy)
    dx, dy = best_shift
    score = (
        0.40 * _block_ssim(a.alpha, _shift_zero(b.alpha, dx, dy))
        + 0.27 * _binary_iou(a.alpha, _shift_zero(b.alpha, dx, dy))
        + 0.20 * _block_ssim(a.gray, _shift_zero(b.gray, dx, dy))
        + 0.08 * _cosine_similarity(a.edge, _shift_zero(b.edge, dx, dy))
        + 0.05 * hash_score
    )
    return max(0.0, min(1.0, score))


def item_icon_path(item: InventoryItem) -> str:
    return clean_text((item.json_obj.get("displayProperties") or {}).get("icon", ""))


def download_icon(
    icon_path: str,
    cache_dir: Path,
    config: IconBuilderConfig,
) -> Tuple[Optional[bytes], str]:
    if not icon_path:
        return None, "manifest_icon_path_empty"
    suffix = Path(icon_path).suffix.lower() or ".png"
    cache_file = cache_dir / f"{hashlib.sha256(icon_path.encode()).hexdigest()[:32]}{suffix}"
    cache_dir.mkdir(parents=True, exist_ok=True)
    if cache_file.exists() and cache_file.stat().st_size > 0:
        return cache_file.read_bytes(), "cache"
    url = (
        icon_path if icon_path.startswith("http")
        else config.bungie_icon_base.rstrip("/") + "/" + icon_path.lstrip("/")
    )
    request = urllib.request.Request(url, headers={"User-Agent": config.http_user_agent})
    try:
        with urllib.request.urlopen(request, timeout=config.http_timeout_seconds) as response:
            data = response.read()
    except (urllib.error.URLError, TimeoutError, OSError) as error:
        return None, f"download_failed:{type(error).__name__}"
    if not data:
        return None, "download_empty"
    cache_file.write_bytes(data)
    return data, "download"


def is_enhanced(item: InventoryItem) -> bool:
    text = " ".join((
        item.name, item.item_type_display, item.item_type_and_tier_display, item.tier_type_name
    )).lower()
    return "强化" in text or "enhanced" in text


def is_normal_weapon_trait(item: InventoryItem) -> bool:
    if item.item_type != 19 or not item.has_plug or is_enhanced(item):
        return False
    type_display = clean_text(item.item_type_display).lower()
    tier_display = clean_text(item.item_type_and_tier_display).lower()
    if any(value in type_display or value in tier_display for value in ("起源", "origin")):
        return False
    return (
        type_display in {"特性", "trait"}
        or tier_display.endswith("特性")
        or tier_display.endswith("trait")
    )


def build_official_visual_catalog(
    index: ManifestIndex,
    official_cache: Path,
    signature_cache: Dict[str, ImageSignature],
    config: IconBuilderConfig,
) -> Tuple[List[OfficialVisual], Dict[int, str], List[Dict[str, Any]]]:
    by_icon_path: Dict[str, List[InventoryItem]] = defaultdict(list)
    for item in index.items:
        if is_normal_weapon_trait(item) and item_icon_path(item):
            by_icon_path[item_icon_path(item)].append(item)

    grouped: Dict[str, Dict[str, Any]] = {}
    errors = []
    for icon_path, items in sorted(by_icon_path.items()):
        data, fetch_method = download_icon(icon_path, official_cache, config)
        if data is None:
            errors.append({
                "icon_path": icon_path,
                "item_hashes": " / ".join(str(item.hash) for item in items),
                "item_names": " / ".join(sorted({item.name for item in items})),
                "reason": fetch_method,
            })
            continue
        cache_key = f"official:{sha256_bytes(data)}"
        signature = signature_cache.get(cache_key)
        if signature is None:
            try:
                signature = signature_from_bytes(data, config)
            except Exception as error:
                errors.append({
                    "icon_path": icon_path,
                    "item_hashes": " / ".join(str(item.hash) for item in items),
                    "item_names": " / ".join(sorted({item.name for item in items})),
                    "reason": f"invalid_image:{type(error).__name__}",
                })
                continue
            signature_cache[cache_key] = signature
        bucket = grouped.setdefault(signature.canonical_sha256, {
            "signature": signature, "items": [], "icon_paths": [], "file_sha256s": [],
        })
        bucket["items"].extend(items)
        bucket["icon_paths"].append(icon_path)
        bucket["file_sha256s"].append(signature.file_sha256)

    visuals = []
    item_visual_map = {}
    for canonical_sha, bucket in sorted(grouped.items()):
        unique_items = {item.hash: item for item in bucket["items"]}
        visual = OfficialVisual(
            visual_id=canonical_sha[:24],
            canonical_sha256=canonical_sha,
            signature=bucket["signature"],
            items=sorted(unique_items.values(), key=lambda item: (item.name, item.hash)),
            icon_paths=sorted(set(bucket["icon_paths"])),
            file_sha256s=sorted(set(bucket["file_sha256s"])),
        )
        visuals.append(visual)
        for item in visual.items:
            item_visual_map[item.hash] = visual.visual_id
    return visuals, item_visual_map, errors


def resolve_global_icons(
    contexts: Sequence[IconContext],
    visuals: Sequence[OfficialVisual],
    signature_cache: Dict[str, ImageSignature],
    config: IconBuilderConfig,
) -> Dict[str, GlobalIconResolution]:
    unique_contexts = {}
    counts: Dict[str, int] = defaultdict(int)
    for context in contexts:
        unique_contexts.setdefault(context.icon_sha256, context)
        counts[context.icon_sha256] += 1

    resolutions = {}
    for position, (icon_sha, context) in enumerate(sorted(unique_contexts.items()), start=1):
        source_key = f"source:{icon_sha}"
        source_signature = signature_cache.get(source_key)
        if source_signature is None:
            source_signature = signature_from_bytes(context.icon_bytes, config)
            signature_cache[source_key] = source_signature
        override_name = clean_text(config.icon_name_overrides.get(icon_sha, ""))
        eligible = [
            visual for visual in visuals
            if not override_name
            or any(norm_name(name) == norm_name(override_name) for name in visual.names)
        ]
        exact = []
        for visual in eligible:
            if source_signature.file_sha256 in visual.file_sha256s:
                exact.append((1.0, visual, "exact_file"))
            elif source_signature.canonical_sha256 == visual.canonical_sha256:
                exact.append((1.0, visual, "exact_canonical_pixel"))
        if exact:
            scored = sorted(exact, key=lambda value: value[1].visual_id)
        else:
            coarse = sorted(
                ((fast_image_similarity(source_signature, visual.signature), visual) for visual in eligible),
                key=lambda value: (-value[0], value[1].visual_id),
            )
            scored = [
                (image_similarity(source_signature, visual.signature, config), visual, "shape_ssim_iou")
                for _, visual in coarse[:config.global_expensive_top_k]
            ]
            scored.sort(key=lambda value: (-value[0], value[1].visual_id))
        best = scored[0] if scored else None
        second_score = 0.0
        if best is not None:
            best_names = {norm_name(name) for name in best[1].names}
            for candidate_score, candidate_visual, _ in scored[1:]:
                if best_names.isdisjoint({norm_name(name) for name in candidate_visual.names}):
                    second_score = candidate_score
                    break
        accepted = False
        if best is None:
            reason = "official_visual_catalog_empty_or_override_not_found"
        elif override_name:
            accepted, reason = True, "accepted_manual_name_override"
        elif best[2] in {"exact_file", "exact_canonical_pixel"}:
            accepted, reason = True, "accepted_exact_global_visual"
        elif not config.allow_approximate_match:
            reason = "approximate_disabled"
        elif best[0] < config.global_min_similarity:
            reason = "global_similarity_below_threshold"
        elif best[0] - second_score < config.global_min_score_margin:
            reason = "global_top_two_visuals_too_close"
        else:
            accepted, reason = True, "accepted_global_shape_match"

        resolutions[icon_sha] = GlobalIconResolution(
            icon_sha256=icon_sha,
            accepted=accepted,
            reason=reason,
            best_visual_id=best[1].visual_id if best else "",
            best_score=round(best[0], 6) if best else 0.0,
            second_score=round(second_score, 6) if best else 0.0,
            margin=round(best[0] - second_score, 6) if best else 0.0,
            match_method=best[2] if best else "",
            candidate_summary=[{
                "visual_id": visual.visual_id,
                "names": visual.names,
                "hashes": visual.hashes,
                "score": round(score, 6),
                "method": method,
                "icon_paths": visual.icon_paths,
            } for score, visual, method in scored[:config.global_top_k]],
            occurrence_count=counts[icon_sha],
        )
        if position % 25 == 0 or position == len(unique_contexts):
            print(f"[全局图标] {position}/{len(unique_contexts)}")
    return resolutions
