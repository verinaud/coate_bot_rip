import winreg
from time import sleep

class PDFSiape():
    
    def impressora_corrente():
        """
        Retorna a impressora atualmente configurada.
        Importante para voltar a configuração da impressora anterior.

        Returns:
            str : nome da impressora atual.
        """
        key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, "Device")
            return value.split(",")[0]
    
    def impressora_padrao(nome_impressora) :
        """
        Retrona à impressora padrão
        
        Args:
            str : nome_impressora
        """        
        try:
            # Abrir a chave do registro onde as configurações da impressora padrão são armazenadas
            key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
                # Definir o nome da impressora PDF como a impressora padrão
                winreg.SetValueEx(key, "Device", 0, winreg.REG_SZ, nome_impressora)
            print(f"Configuração da impressora padrão para '{nome_impressora}' foi atualizada com sucesso.")
        except Exception as e:
            print(f"Erro ao definir a impressora padrão: {e}")

    def print_to_pdf() :
        """
        Modifica a impressora atual do Windows para "Microsoft Print to PDF".
        
        """
        nome_impressora = "Microsoft Print to PDF"
        try:
            # Abrir a chave do registro onde as configurações da impressora padrão são armazenadas
            key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
                # Definir o nome da impressora PDF como a impressora padrão
                winreg.SetValueEx(key, "Device", 0, winreg.REG_SZ, nome_impressora)
            print(f"Configuração da impressora padrão para '{nome_impressora}' foi atualizada com sucesso.")
        except Exception as e:
            print(f"Erro ao definir a impressora padrão: {e}")

    def coletar_tela(self) :
        """
        Coleta a tela siape e guarda na memória.
        
        """
        self.dlg.type_keys('%a')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)

    def salva_telas_coletadas(self) :
        """
        Salva as telas siape coletadas num arquivo PDF e salva no computador.
        """
        self.dlg.type_keys('%a')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%i')

    def limpar_memoria(self) :
        """
        Exclui as telas coletadas da memória.
        """
        self.dlg.type_keys('%a')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys('%c')
        sleep(0.1)
        self.dlg.type_keys("{RIGHT}")
        sleep(0.1)
        self.dlg.type_keys('%e')