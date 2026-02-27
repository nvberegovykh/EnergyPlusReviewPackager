#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Launcher: opens EnergyPlusReviewPackager without a terminal (uses pythonw on Windows)."""
import os
import runpy
_here = os.path.dirname(os.path.abspath(__file__))
_main = os.path.join(_here, "EnergyPlusReviewPackager.py")
runpy.run_path(_main, run_name="__main__")
