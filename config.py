#!/usr/bin/env python3

import os


class Config:
    project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    cfg_file = project_dir + "/config.json"
