# 📸 Photobooth Business Lead & ROI Calculator Bot

A professional Telegram bot built with **Aiogram 3.x** designed to automate lead generation and financial forecasting for photobooth businesses. The bot guides users through an interactive survey, calculates potential revenue, and provides instant reports to administrators.

## 🚀 Key Features

- **Interactive ROI Calculator**:
  - Calculates three revenue scenarios (Conservative 7%, Realistic 10%, Potential 15%) based on monthly foot traffic.
  - Prices are dynamically calculated using configurable rates.
- **Smart Lead Generation (FSM)**:
  - Multi-step survey using Finite State Machine to ensure data integrity.
  - Validation using Regular Expressions (protects against invalid characters/emojis).
- **Professional Excel CRM**:
  - Automatically saves every lead to a stylized `.xlsx` file.
  - **Date** column is prioritized (first column).
  - Fixed light-gray header row with bold text.
  - Automatic column width adjustment for better readability.
- **Multi-Admin Support**:
  - Real-time notifications sent to multiple `ADMIN_IDS`.
  - Secure "Download Database" button available only for verified admins.
- **Marketing Integration**:
  - Automated PDF catalog delivery.
  - Direct "Contact Manager" buttons.

## 📂 Project Structure

```text
├── assets/             # PDF catalogs and media files
├── db/                 # Local storage for Excel leads database
├── .env                # Environment variables (excluded from Git)
├── .gitignore          # List of ignored files and folders
├── main.py             # Main bot logic
├── bot_log.log         # Auto-generated log file
├── requirements.txt    # Project dependencies
└── README.md           # Documentation
```

## 🛠 Tech Stack

- Python 3.9+
- Aiogram 3.x: Telegram Bot API framework.
- Pandas & Openpyxl: Professional Excel report generation and styling.
- Python-dotenv: Secure environment variable management.

## ⚙️ Installation & Setup

Clone the repository:

```bash
git clone https://github.com/ksalab/photobooth-bot.git
cd photobooth-bot
```

Create and activate a virtual environment:

```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Configure Environment Variables:

Create a .env file in the root directory:

```env
BOT_TOKEN=your_telegram_bot_token
ADMIN_IDS=12345678,98765432
MANAGER_URL=https://t.me/your_username
DB_PATH=db/leads.xlsx
FILE_PATH=assets/catalog.pdf
```

Run the bot:

```bash
python main.py
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
