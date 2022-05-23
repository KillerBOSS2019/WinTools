import soundmeter
from math import *

def root_mean_square():
    current_rms = 0
    # DO SOME STUFF
    return current_rms


def decibel_a():
    current_dba = log(20, 10) * (root_mean_square() / 20)
    return current_dba

print(decibel_a())