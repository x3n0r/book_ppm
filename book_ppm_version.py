version="V1.0.2"

#function to check if version is greater or lower
def versiontuple(v):
    v = v[1:]
    return tuple(map(int, (v.split("."))))