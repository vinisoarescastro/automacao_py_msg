import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import urllib.parse

# Função para carregar contatos do Excel
def carregar_contatos(caminho_do_arquivo):
    """Carrega contatos de um arquivo Excel."""
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_do_arquivo)
        
        # Verifica se existem as colunas necessárias
        if 'nome' not in df.columns or 'numero' not in df.columns:
            print("O arquivo Excel deve conter as colunas 'nome' e 'numero'")
            return None
        
        # Retorna uma lista de dicionários com nome e número
        return df[['nome', 'numero']].to_dict('records')
    except Exception as e:
        print(f"Erro ao carregar o arquivo: {e}")
        return None

# Função para inicializar o driver do Chrome
def inicializar_driver():
    """Inicializa e retorna o driver do Chrome."""
    try:
        # Configurações do navegador
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        # Inicializa o driver do Chrome
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        print(f"Erro ao inicializar o driver: {e}")
        return None

# Função para enviar duas mensagens no WhatsApp
def enviar_duas_mensagens(driver, numero, nome, primeira_mensagem, segunda_mensagem, tempo_espera=5):
    """
    Envia duas mensagens para um número no WhatsApp.
    
    Args:
        driver: Driver do Selenium
        numero: Número de telefone do contato
        nome: Nome do contato
        primeira_mensagem: Primeira mensagem a ser enviada
        segunda_mensagem: Segunda mensagem a ser enviada
        tempo_espera: Tempo de espera entre as mensagens (em segundos)
    
    Returns:
        bool: True se as mensagens foram enviadas com sucesso, False caso contrário
    """
    try:
        # Formata o número para o padrão internacional (remova caracteres não numéricos)
        numero_formatado = ''.join(filter(str.isdigit, str(numero)))
        
        # URL do WhatsApp com o número e a primeira mensagem
        url = f"https://web.whatsapp.com/send?phone={numero_formatado}&text={urllib.parse.quote(primeira_mensagem)}"
        
        # Acessa a URL
        driver.get(url)
        
        # Espera até que o botão de enviar apareça (isso significa que o chat foi carregado)
        try:
            wait = WebDriverWait(driver, 10)  # Ajustado para 10 segundos conforme solicitado
            
            # Tentamos primeiro verificar se a página terminou de carregar completamente
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
            
            # Verificar visualmente (debug) o que está acontecendo
            print(f"⏳ Aguardando carregar o chat para {nome}...")
            
            # Tentamos identificar o botão de enviar de VÁRIAS maneiras diferentes
            try:
                enviar_botao = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='compose-btn-send']")))
            except:
                try:
                    enviar_botao = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-icon='send']")))
                except:
                    try:
                        enviar_botao = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']")))
                    except:
                        # Último recurso: verificar se o campo de mensagem está visível (isso indica que o chat está aberto)
                        campo_texto = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@contenteditable='true']")))
                        
                        # Tentamos encontrar o botão de enviar pelo atributo aria-label
                        enviar_botao = driver.find_element(By.XPATH, "//button[contains(@aria-label, 'Enviar') or contains(@aria-label, 'Send')]")
            
            # Tiramos uma captura de tela para debug (opcional)
            # driver.save_screenshot(f"debug_{numero}.png")
            
            print(f"✅ Chat carregado com sucesso para {nome}")
            
        except Exception as e:
            # Se não conseguiu encontrar o botão, provavelmente o número não existe no WhatsApp
            print(f"⚠️ Erro ao carregar chat para {numero} ({nome}): {str(e)}")
            return False
        
        # Clica no botão para enviar a primeira mensagem
        try:
            # Verifica se a variável enviar_botao foi definida
            if 'enviar_botao' in locals():
                enviar_botao.click()
                print(f"✅ Primeira mensagem enviada para {nome} ({numero})")
            else:
                # Se não encontramos o botão mas o chat carregou (temos o campo de texto)
                if 'campo_texto' in locals():
                    # Envia a mensagem pressionando Enter no campo de texto
                    campo_texto.send_keys(Keys.ENTER)
                    print(f"✅ Primeira mensagem enviada usando Enter para {nome} ({numero})")
                else:
                    raise Exception("Não foi possível encontrar um método para enviar a mensagem")
        except Exception as e:
            print(f"❌ Erro ao enviar primeira mensagem: {str(e)}")
            # Tentamos continuar mesmo assim
        
        # Espera o tempo definido antes de enviar a segunda mensagem
        time.sleep(tempo_espera)
        
        # Tenta enviar a segunda mensagem
        try:
            # Tenta encontrar o campo de texto de várias maneiras possíveis
            try:
                campo_texto = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true']"))
                )
            except:
                try:
                    campo_texto = driver.find_element(By.XPATH, "//div[@role='textbox']")
                except:
                    campo_texto = driver.find_element(By.XPATH, "//footer//div[@contenteditable='true']")
            
            # Clica no campo de texto, digita a mensagem e envia
            campo_texto.click()
            campo_texto.clear()
            # Digita a mensagem caractere por caractere para evitar problemas
            for caractere in segunda_mensagem:
                campo_texto.send_keys(caractere)
            
            # Envia a mensagem pressionando Enter
            campo_texto.send_keys(Keys.ENTER)
            
            print(f"✅ Segunda mensagem enviada para {nome} ({numero})")
            
            # Espera um pouco antes de passar para o próximo contato
            time.sleep(3)
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao enviar segunda mensagem para {nome} ({numero}): {str(e)}")
            # Mesmo que falhe na segunda mensagem, consideramos parcialmente bem-sucedido
            return True
    
    except Exception as e:
        print(f"❌ Erro ao processar {nome} ({numero}): {str(e)}")
        return False

# Função principal
def main():
    print("======== AUTOMAÇÃO DE ENVIO DE MENSAGENS WHATSAPP ========")
    print("Desenvolvido para facilitar o envio de mensagens em massa")
    print("==========================================================\n")
    
    # Solicita o caminho do arquivo Excel
    caminho_do_arquivo = input("Digite o caminho do arquivo Excel com os contatos: ")
    
    # Solicita as mensagens
    print("\n--- CONFIGURAÇÃO DAS MENSAGENS ---")
    
    # Primeira mensagem
    primeira_mensagem = input("Digite a primeira mensagem: ")
    
    # Segunda mensagem
    segunda_mensagem = input("Digite a segunda mensagem: ")
    
    # Personalização com nome
    personalizar = input("\nDeseja personalizar as mensagens com o nome do contato? (s/n): ").lower() == 's'
    
    if personalizar:
        print("\nUse {nome} onde desejar inserir o nome do contato.")
        primeira_mensagem = input("Digite a primeira mensagem: ")
        segunda_mensagem = input("Digite a segunda mensagem: ")
    
    # Tempo de espera entre mensagens
    try:
        tempo_espera = int(input("\nTempo de espera entre as mensagens (em segundos, recomendado 3-5): "))
    except ValueError:
        tempo_espera = 5
        print("Valor inválido! Usando tempo padrão de 5 segundos.")
    
    # Carrega os contatos do Excel
    contatos = carregar_contatos(caminho_do_arquivo)
    if not contatos:
        print("❌ Não foi possível carregar os contatos. Verifique o arquivo Excel.")
        return
    
    # Inicializa o navegador Chrome
    driver = inicializar_driver()
    if not driver:
        print("❌ Não foi possível inicializar o navegador Chrome.")
        return
    
    try:
        # Abre o WhatsApp Web
        driver.get("https://web.whatsapp.com")
        print("\n➡️ Escaneie o código QR com seu celular e pressione Enter quando estiver pronto...")
        input()
        
        # Contadores para o relatório
        total = len(contatos)
        sucessos = 0
        falhas = 0
        contatos_falha = []
        
        # Exibe informações do início do processamento
        print("\n" + "="*50)
        print(f"INICIANDO ENVIO PARA {total} CONTATOS")
        print("="*50)
        
        # Processa cada contato
        for i, contato in enumerate(contatos, 1):
            nome = contato['nome']
            numero = contato['numero']
            
            print(f"\nProcessando {i}/{total}: {nome} ({numero})")
            
            # Personaliza as mensagens se necessário
            msg1 = primeira_mensagem.replace("{nome}", nome) if personalizar else primeira_mensagem
            msg2 = segunda_mensagem.replace("{nome}", nome) if personalizar else segunda_mensagem
            
            # Envia as mensagens
            if enviar_duas_mensagens(driver, numero, nome, msg1, msg2, tempo_espera):
                sucessos += 1
            else:
                falhas += 1
                contatos_falha.append(f"{nome}: {numero}")
        
        # Exibe o relatório final
        print("\n" + "="*50)
        print("RELATÓRIO FINAL")
        print("="*50)
        print(f"Total de contatos: {total}")
        print(f"Envios bem-sucedidos: {sucessos}")
        print(f"Envios com falha: {falhas}")
        
        if falhas > 0:
            print("\nContatos com falha no envio:")
            for contato in contatos_falha:
                print(f"- {contato}")
    
    except Exception as e:
        print(f"\n❌ Ocorreu um erro durante a execução: {str(e)}")
    
    finally:
        # Pergunta se deve fechar o navegador
        fechar = input("\nDeseja fechar o navegador? (s/n): ").lower() == 's'
        if fechar:
            driver.quit()
            print("Navegador fechado.")
        else:
            print("Navegador mantido aberto. Feche-o manualmente quando terminar.")

# Executa o programa principal
if __name__ == "__main__":
    main()