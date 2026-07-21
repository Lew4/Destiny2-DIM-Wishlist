#!/usr/bin/env python3
"""Backward-compatible entry point for the modular DIM wishlist builder."""

from dim_wishlist.cli import main


if __name__ == "__main__":
    raise SystemExit(main())
