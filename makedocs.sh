#!/bin/bash

set -euo pipefail

project_dir="$( cd "$(dirname "$0")" >/dev/null 2>&1 ; pwd -P )"
echo "generating documents for $project_dir"
cd /tmp
[[ -e adrdox ]] && rm -rf adrdox
git clone --depth=1 https://github.com/adamdruppe/adrdox
cp "$project_dir"/.skeleton.html adrdox/skeleton.html
cd adrdox
make
./doc2 -i "$project_dir"/source
mv generated-docs/* "$project_dir"/docs
cp "$project_dir"/docs/xlld.html "$project_dir"/docs/index.html
