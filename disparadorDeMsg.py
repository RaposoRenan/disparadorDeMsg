import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import threading

def enviar_mensagem_whatsapp(nome, telefone, mensagem):
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(30)
        seta = pyautogui.locateCenterOnScreen('Capturar.PNG')
        sleep(3)
        pyautogui.click(seta[0], seta[1])
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(3)
        return True
    except pyautogui.ImageNotFoundException:
        print(f'Não foi possível encontrar a imagem para clicar em {nome}')
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
        return False

def main():
    # Ler planilha e guardar informações sobre nome, telefone e data de vencimento
    workbook = openpyxl.load_workbook('arqvFormat.xlsx')
    pagina_clientes = workbook['Sheet1']

    continuar_execucao = True
    paused = False

    # Definir mensagem padrão
    mensagem_padrao = '''
Decisões, decisões... Entre 2,00 ou um presente misterioso, qual será a escolha?
Venha descobrir os detalhes do nosso primeiro dia de aula na UNIP São José do Rio Preto!

#PrimeiroDiaDeAula #UNIPrioPreto

https://www.instagram.com/reel/C3gMXapuz9B/?igsh=a2IwZXVkZjk1YzZl
    '''

    def worker():
        nonlocal paused
        for linha in pagina_clientes.iter_rows(min_row=2):
            if paused:
                print("Programa pausado. Pressione 'c' para continuar ou 'q' para sair.")
                while paused:
                    sleep(1)
                print("Continuando a execução...")
            nome = linha[0].value
            telefone = linha[1].value

            if nome is None or telefone is None:
                print("Encontrado valor None. Encerrando o programa.")
                break

            sucesso = enviar_mensagem_whatsapp(nome, telefone, mensagem_padrao)
            if not sucesso:
                print("Programa pausado. Digite 'c' para continuar ou 'q' para sair.")
                while paused:
                    sleep(1)
                print("Continuando a execução...")
            if not continuar_execucao:
                break

    thread = threading.Thread(target=worker)
    thread.start()

    while continuar_execucao:
        if not paused:
            opcao = input("Digite 'p' para pausar: ").lower()
            if opcao == 'p':
                paused = True
        else:
            opcao = input("Programa pausado. Digite 'c' para continuar ou 'q' para sair: ").lower()
            if opcao == 'c':
                paused = False
            elif opcao == 'q':
                continuar_execucao = False
                paused = False
            else:
                print("Opção inválida. Por favor, digite 'c' ou 'q'.")

    thread.join()
    print("Programa encerrado.")

if __name__ == "__main__":
    main()
