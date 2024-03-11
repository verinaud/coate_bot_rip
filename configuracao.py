import json

class Config:
    def __init__(self):
        self.config = None
        self.popula_config()
        
    def get_config(self):
        return self.config
    
    def set_config(self, config):
        self.config = config
        with open("config.json", 'w', encoding="utf-8") as file:
            json.dump(config, file, indent=4, ensure_ascii=False)
    
    def popula_config(self):
        try:
            with open("config.json", encoding="utf-8") as file:
                self.config = json.load(file)
        except FileNotFoundError:
            self.default()
    
    #andluve@gmail.com  
    
    def default(self):
        config_json = {
            "ultimo_acesso_user": "",
            "senha": "",
            "ultimo_orgao": "MGI",
            "url_sei": "https://sei.economia.gov.br",
            "url_siapenet": "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio",
            "unidade_sei": "MGI-SGP-DECIPEX-COATE-CADAS",
            "marcador": "CDCONVINC",
            "funcionando_teste": False,
            "diretorio_vinculos" : "C://PDF",
            "diretorio_log": "C://LogCoate",
            "manual": True,
            "lista_orgaos": [
                "MGI", "ME", "CMB", "COAF", "MTP", "MF", "MPO", "MDIC", "MPI", "MPS", "MEMP"
            ],
            "lista_marcadores_substitutos": [
                "CADASTRO ATIVO", "EXCLUÍDO BATIMENTO", "EXCLUÍDO FALECIMENTO", "VERIFICAÇÃO MANUAL", "SEM VÍNCULO"
            ],
            "vinculos_decipex": ["40802", "40805", "40806"],
            "lista_ativo": ["", "00/000"],
            "lista_excluido_batimento": ["02/227", "02/237", "07/132"],
            "lista_excluido_falecimento": [
                "02/007", "02/073", "02/101", "02/118", "02/128", "02/187", "02/193", "02/207", "02/209", "02/215",
                "02/224", "02/225", "07/003", "02/301", "02/302", "02/410", "02/411", "02/216"
            ]
        }
        
        with open("config.json", 'w', encoding="utf-8") as file:
            json.dump(config_json, file, indent=4, ensure_ascii=False)
            
        with open("config.json", encoding="utf-8") as file:
                self.config = json.load(file)
