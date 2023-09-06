version="V1.0.1"

#function to check if version is greater or lower
def versiontuple(v):
    v = v[1:]
    return tuple(map(int, (v.split("."))))