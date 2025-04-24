import os
import keyboard
import funcoes

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def criar_excel():
    api = input("API que desejas importar: ")
    nome = input("Nome do ficheiro: ")

    funcoes.criar_excel(f"https://transparencia.sns.gov.pt/api/explore/v2.1/catalog/datasets/{api}/records?limit=100", nome)

def atualizar_excel():
    api = input("API que desejas importar: ")
    ficheiro = input("Ficheiro que queres fazer a atualização: ")
    funcoes.atualizar_excel(f"https://transparencia.sns.gov.pt/api/explore/v2.1/catalog/datasets/{api}/records?limit=100", ficheiro)

def menu():
    options = ["1 - Criar novo Excel", "2 - Atualizar um Excel", "0 - Sair"]

    while True:
        clear_screen()
        print("=== Menu ===")
        for option in options:
            print(f"{option}")
        key = int(input(f"\nEscolha a sua opção: "))

        if key == 0:
            return
        elif key == 1:
            clear_screen()
            criar_excel()
        elif key == 2:
            clear_screen()
            atualizar_excel()        

if __name__ == "__main__":
    menu()
