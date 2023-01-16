#!/usr/bin/env bash
set -eu -o pipefail

cd "$(git rev-parse --show-toplevel)"
poetry run flake8 .
poetry run mypy --config-file=pyproject.toml .
