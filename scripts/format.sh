#!/usr/bin/env bash
set -eu -o pipefail

cd "$(git rev-parse --show-toplevel)"
poetry run black --config=pyproject.toml .
poetry run isort --settings-file=pyproject.toml .
