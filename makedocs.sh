set -e -o pipefail
kal_project_dir=`pwd`
echo generating documents for ${kal_project_dir}
cd ~
mkdir -p tmp
cd tmp
rm -rf adrdox
git clone https://github.com/adamdruppe/adrdox
cp ${kal_project_dir}/.skeleton.html adrdox/skeleton.html
cd adrdox
make
./doc2 -i ${kal_project_dir}/source
mv generated-docs/* ${kal_project_dir}/docs
cp ${kal_project_dir}/docs/xlld.html ${kal_project_dir}/docs/index.html
cd ${kal_project_dir}
echo succeeded - docs generated
