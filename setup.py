from setuptools import Command, setup, find_packages


# see StackOverflow/458550
exec(open('src/pycel/version.py').read())


setup(
    name='Pycel',
    version=__version__,
    packages=find_packages('src'),
    package_dir={'': 'src'},
    description='A library for compiling MS Excel spreadsheets to Python & visualizing them as a graph',
    url='https://github.com/dgorissen/pycel',
    install_requires=open('requirements.txt', 'r').read().splitlines(),
    tests_require=['pytest-cov>=2.5.1,<3.0'],
    test_suite='nose.collector',
    author='Dirk Gorissen',
    author_email='dgorissen@gmail.com',
    long_description=''.join(list(open('README.rst', 'r'))[3:6]),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: Implementation :: PyPy',
    ],
    include_package_data=True,
    zip_safe=True)
