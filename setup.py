from setuptools import setup

setup(name='music-cue',
      version='0.1',
      description='A data entry application',
      url='https://github.com/hans-vvv/Music-Cue',
      author='Hans Verkerk',
      author_email='verkerk.hans@gmail.com',
      license='MIT',
      packages=['music_cue'],
      entry_points={'console_scripts': ['music_cue=music_cue.command_line:main']},
      zip_safe=False)
