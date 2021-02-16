from fuzzywuzzy import fuzz
from fuzzywuzzy import process



if __name__ == '__main__':
    a = fuzz.ratio('Привет мир', 'Привт rир')
    print(a)