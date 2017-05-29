from setuptools import Command, setup, find_packages


# see StackOverflow/458550
exec(open('src/pycel/version.py').read())


setup(name='Pycel',
      version=__version__,
      packages=find_packages('src'),
      package_dir={'': 'src'},
      description='A library for compiling MS Excel spreadsheets to Python & visualizing them as a graph',
      url='https://github.com/dgorissen/pycel',
      tests_require=['nose>=1.2'],
      test_suite='nose.collector',
      install_requires=open('requirements.txt', 'r').read().splitlines(),
      author='Dirk Gorissen',
      author_email='dgorissen@gmail.com',
      long_description=''.join(list(open('README.md', 'r'))[3:6]),
      classifiers=['Development Status :: 4 - Beta',
                   'Intended Audience :: Developers',
                   'License ::  OSI Approved ',
                   ]
      )
