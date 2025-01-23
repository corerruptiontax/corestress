python

     from setuptools import setup, find_packages  
  
     setup(  
         name='corestress',  
         version='0.1',  
         packages=find_packages(),  
         install_requires=[  
             'pandas',  
             'requests',  
         ],  
     )  
     ```
