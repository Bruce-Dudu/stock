# -*- coding: utf-8 -*-
"""
Created on Thu May 31 10:45:35 2018

@author: Adu
"""
import os
import win32com.client
from win32com.client import *
import pandas as pd
import numpy as np


# 创建DataFrame对象
df = pd.DataFrame([1, 2, 3, 4, 5], columns=['cols'], index=['a','b','c','d','e'])
print("DF ")
print (df)
