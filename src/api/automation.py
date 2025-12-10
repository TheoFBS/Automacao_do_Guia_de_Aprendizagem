import logging
from .google_sheets import Sheets
from .word import Word
from config.settings import (
    ID_FUNDAMENTAL_I,
    ID_FUNDAMENTAL_II,
    ID_MEDIO
)

logger = logging.getLogger(__name__)

class Automation:
    def __init__(self, crendentials, bot_credentials, token, scopes):
        self.spreadsheet = Sheets(crendentials, bot_credentials, token, scopes)
        self.sheet_id = None
        self.sheet_range = None
        self.serie = None
        self.bimestre = None
    
    def prepare_data(self, data: dict):
        self.bimestre = data['bimestre']
        self.serie = serie = data['serie']
        disciplina = data['disciplina']
        ss = None
        
        if serie in ['1° ano','2° ano','3° ano','4° ano','5° ano','1º ano','2º ano','3º ano','4º ano','5º ano']:
            self.sheet_id = ID_FUNDAMENTAL_I
            ss = 1
        elif serie in ['6° ano','7° ano','8° ano','9° ano','6º ano','7º ano','8º ano','9º ano']:
            self.sheet_id = ID_FUNDAMENTAL_II
            ss = 2
        else:
            self.sheet_id = ID_MEDIO
            ss = 3
        logger.info(f"Alterando planilha {ss}")
        
        match disciplina:
            case "Arte":
                if ss == 1: sheet_range = "A2:J322"
                if ss == 2: sheet_range = "A2:J170"
                if ss == 3: sheet_range = "A1:J57"
            case "Arte e Mídias Digitais":
                sheet_range = "A1:J57"
            case "Biologia":
                sheet_range = "A1:J113"
            case "Biotecnologia":
                sheet_range = "A1:L57"
            case "Ciências":
                if ss == 1: sheet_range = "A2:J162"
                if ss == 2: sheet_range = "A2:J394"
            case "Educação Financeira":
                if ss == 2: sheet_range = "A2:I170"
                if ss == 3: sheet_range = "A1:I113"
            case "Educação Física":
                if ss == 1: sheet_range = "A2:J322"
                if ss == 2: sheet_range = "A2:K226"
                if ss == 3: sheet_range = "A1:K113"
            case "Empreendedorismo":
                sheet_range = "A1:J85"
            case "Filosofia":
                sheet_range = "A1:I58"
            case "Filosofia e Sociedade Moderna":
                sheet_range = "A1:J57"
            case "Física":
                sheet_range = "A1:I169"
            case "Geografia":
                if ss == 1: sheet_range = "A2:J162"
                if ss == 2: sheet_range = "A2:J282"
                if ss == 3: sheet_range = "A1:I113"
            case "Geopolítica":
                sheet_range = "A1:M57"
            case "História":
                if ss == 1: sheet_range = "A2:J162"
                if ss == 2: sheet_range = "A2:J282"
                if ss == 3: sheet_range = "A1:I169"
            case "Liderança":
                sheet_range = "A1:M57"
            case "Língua Inglesa":
                if ss == 1: sheet_range = "A2:J162"
                if ss == 2: sheet_range = "A2:J298"
                if ss == 3: sheet_range = "A2:J242"
            case "Língua Portuguesa":
                if ss == 1: sheet_range = "A2:K770"
                if ss == 2: sheet_range = "A2:L450"
                if ss == 3: sheet_range = "A1:L62"
            case "Matemática":
                if ss == 1: sheet_range = "A2:J790"
                if ss == 2: sheet_range = "A2:J562"
                if ss == 3: sheet_range = "A1:J54"
            case "Oratória":
                sheet_range = "A1:J85"
            case "Programação":
                sheet_range = "A1:P62"
            case "Projeto de Convivência":
                if ss == 1: sheet_range = "A2:H188"
            case "Projeto de Vida":
                if ss == 2: sheet_range = "A2:M114"
                if ss == 3: sheet_range = "A1:K85"
            case "Química":
                sheet_range = "A1:I113"
            case "Química Aplicada":
                sheet_range = "A1:J57"
            case "Redação e Leitura":
                if ss == 2: sheet_range = "A2:K257"
                if ss == 3: sheet_range = "A1:K196"
            case "Robótica - PEI 9h":
                if ss == 2: sheet_range = "A2:K86"
                if ss == 3: sheet_range = "A1:K29"
            case "Sociologia":
                sheet_range = "A1:I58"
            case "Tecnologia e Inovação":
                sheet_range = "A2:P261"
                
            case "Tecnologia e Inovação PEI 7h e 9h":
                if ss == 2: sheet_range = "A2:P261"
                sheet_range = "A1:J57"
        self.sheet_range = disciplina + "!" + sheet_range
        logger.info(f"Bimestre: {self.bimestre}, Serie: {self.serie}, Range: {self.sheet_range}")
    
    def process_document(self, template_path: str, output_path: str,
                         dados: dict):
        self.prepare_data(dados)
        
        if type(self.sheet_range) == []:
            data = self.spreadsheet.batch_get_values(self.sheet_id, self.sheet_range)
        else:
            data = self.spreadsheet.get_values(self.sheet_id, self.sheet_range)
        logger.info("Conteudo salvo com sucesso")
        
        process = Word(template_path)
        process.find_all_placeholders()
        process.fill_table_GA(data, self.serie, self.bimestre)
        process.replace_placeholders(dados)
        process.save_document(output_path)