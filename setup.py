from setuptools import setup, find_packages


def read(filename):
    with open(filename, 'r', encoding='utf-8') as myfile:
        return myfile.read()


setup(name='template-pptx-jinja',
    version="0.1.0",
    description='PowerPoint presentation builder from template using Jinja2',
    long_description=read('readme.md'),
    long_description_content_type='text/markdown',
    url='https://github.com/Thykof/template-pptx-jinja',
    author='Thykof',
    author_email='thykof@protonmail.ch',
    install_requires=read('requirements.txt').split(),
    license='',
    packages=find_packages(),
    keywords=['powerpoint', 'ppt', 'pptx', 'template'],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Topic :: Utilities",
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent"
    ]
)
