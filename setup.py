# coding=utf-8
import setuptools

with open('README.rst') as f:
    README = f.read()

setuptools.setup(
    name='xcelios',
    version='0.0.1',
    author='ThetaDev',
    description='OpenPyXL Excel templating tool',
    long_description=README,
    long_description_content_type='text/x-rst',
    license='MIT License',
    url="https://github.com/Theta-Dev/xcelios",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    install_requires=[
        'openpyxl'
    ],
    packages=setuptools.find_packages(exclude=['tests*']),
)