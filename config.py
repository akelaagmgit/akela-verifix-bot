"""Konfiguratsiya - .env dan o'qiladi."""
import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()

@dataclass
class Settings:
    bot_token: str = os.getenv("BOT_TOKEN", "")
    verifix_login: str = os.getenv("VERIFIX_LOGIN", "admin@akela")
    verifix_password: str = os.getenv("VERIFIX_PASSWORD", "")
    verifix_company_code: str = os.getenv("VERIFIX_COMPANY_CODE", "")
    verifix_filial_id: str = os.getenv("VERIFIX_FILIAL_ID", "483143")
    verifix_user_id: str = os.getenv("VERIFIX_USER_ID", "")
    verifix_project_code: str = os.getenv("VERIFIX_PROJECT_CODE", "vhr")
    verifix_base_url: str = os.getenv("VERIFIX_BASE_URL", "https://app.verifix.com")
    template_path: str = os.getenv("TEMPLATE_PATH", "template.xlsx")
    output_dir: str = os.getenv("OUTPUT_DIR", "reports")
    allowed_chat_ids: str = os.getenv("ALLOWED_CHAT_IDS", "")

    @property
    def allowed_ids(self) -> set[int]:
        if not self.allowed_chat_ids.strip():
            return set()
        out = set()
        for part in self.allowed_chat_ids.split(","):
            part = part.strip()
            if part.lstrip("-").isdigit():
                out.add(int(part))
        return out

settings = Settings()
