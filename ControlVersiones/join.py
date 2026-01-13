import os


def merge_archivos():
    # Variables
    directory = os.getcwd()

    # iteramos sobre los .txt
    for filename in os.listdir(directory):
        print(filename)
        if filename.endswith(".txt"):
            with open(directory + filename, encoding="utf8") as fp:
                data = fp.read()

            with open("registro_unico.txt", "w", encoding="utf8") as fp:
                fp.write(data)

    fp.close()

    return


def main():
    print("hola")
    merge_archivos()


if __name__ == "__main__":
    main()
