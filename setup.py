from setuptools import setup, find_packages

setup(
    name="KYC_Viewer",
    version="1.0.0",
    description="KYC Exceptions Generator - Audit Automation Tool",
    author="Your Name",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "openpyxl",
        # Add other dependencies here
    ],
    include_package_data=True,
)
