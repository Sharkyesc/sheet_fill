import os

class Config:
    # OpenAI API configuration
    OPENAI_API_KEY = 'sk-rfCIGhxrzcdsMV4jC17e406bE56c47CbA5416068A62318D3'
    OPENAI_BASE_URL = 'http://ipads.chat.gpt:3006/v1/'
    OPENAI_MODEL = 'gemini-2.5-pro-preview-06-05'

    # File path configuration
    INPUT_DIR = './examples'
    OUTPUT_DIR = './output'
    MID_DIR = './mid_docs'
    TEMP_DIR = './temp'

    # Highlight color configuration
    HIGHLIGHT_COLOR = 'FFFF00'  # Yellow

    @classmethod
    def create_directories(cls):
        for directory in [cls.INPUT_DIR, cls.OUTPUT_DIR, cls.TEMP_DIR, cls.MID_DIR]:
            os.makedirs(directory, exist_ok=True) 