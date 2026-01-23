# -*- coding: utf-8 -*-
import importlib.metadata

packages = importlib.metadata.distributions()

for pkg in packages:
    name = pkg.metadata["Name"]
    version = pkg.version
    print(f"{name} | Version: {version}")
