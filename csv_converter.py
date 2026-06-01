import pandas as pd

def converter(a_path, separator_sign):
    try:
        clean_path = a_path.replace('"', '')
        df = pd.read_csv(clean_path)
        df.to_csv(clean_path.replace('.csv', '_separated.csv'),
                  index=False,
                  sep=separator_sign)
        return 0
    except FileNotFoundError:
        return 1
    except Exception:
        return 2

csv_path = input('enter a path:')
sep_sign = input('enter separator sign:')

if __name__ == '__main__':
    print(converter(csv_path, sep_sign))
