from distutils.core import setup

setup(
    name='redirect-tools',
    version='1.0',
    packages=['redirect-tools'],
    scripts=['redirect-tools.py'],
    license='LICENSE.txt',
    description='Tools for dealing with URL redirects.',
    long_description=open('README.txt').read(),
    install_requires=[
        'pandas>=0.20.1'
        'requests>=2.14.2'
        'configparser>=3.5.0'
        'xlutils>=2.0.0'
        'xlrd>=1.0.0'
        'xlwt>=1.2.0'
    ],
)
