import setuptools

setuptools.setup(
  name='ppt-to-pdf-converter',
  version='0.0.2',
  description='convert ppt files to pdf in folder',
  author='jhheo',
  url='https://github.com/Thongangerge/ppt-to-pdf-converter',
  download_url='https://github.com/Thongangerge/ppt-to-pdf-converter',
  packages=['jhconverter'],
  classifiers=[
    "Programming Language :: Python :: 3",
  ],
  install_requires=[
    'pywin32'
  ]
)
