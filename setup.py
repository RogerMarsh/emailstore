# setup.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)
"""emailstore setup file."""

from setuptools import setup

if __name__ == "__main__":

    long_description = open("README").read()
    install_requires=["solentware-misc==1.3.2"]

    setup(
        name="emailstore",
        version="1.4.4",
        description="Select and store emails from mailbox",
        author="Roger Marsh",
        author_email="roger.marsh@solentware.co.uk",
        url="http://www.solentware.co.uk",
        packages=[
            "emailstore",
            "emailstore.core",
            "emailstore.gui",
            "emailstore.help",
        ],
        package_data={
            "emailstore.help": ["*.txt"],
        },
        long_description=long_description,
        license="BSD",
        classifiers=[
            "License :: OSI Approved :: BSD License",
            "Programming Language :: Python :: 3.6",
            "Programming Language :: Python :: 3.7",
            "Programming Language :: Python :: 3.8",
            "Programming Language :: Python :: 3.9",
            "Operating System :: OS Independent",
            "Topic :: Other/Nonlisted Topic",
            #'Topic :: Text Processing',
            #'Topic :: Communications :: Email',
            #'Topic :: Communications :: Email :: Email Client (MUA)',
            "Intended Audience :: End Users/Desktop",
            "Intended Audience :: Developers",
            "Development Status :: 3 - Alpha",
        ],
        install_requires=install_requires,
        dependency_links=[
            "-".join(required.split("==")).join(
                ("http://solentware.co.uk/files/", ".tar.gz")
            )
            for required in install_requires
        ],
    )
