# KTS Billing Dashboard

PHP dashboard for importing monthly customer master data, updating daily payment status, and preparing invoice messages for WhatsApp.

## Requirements

- PHP 8.1 or later with `mysqli`, `zip`, and `dom` extensions
- MySQL with an existing application database
- Apache or another PHP-capable web server

## Setup

1. Copy `config.example.php` to `config.php`.
2. Set `DB_HOST`, `DB_USER`, `DB_PASS`, and `DB_NAME` for the server. Set Mayar credentials through environment variables if the existing application integrations need them.
3. Place the project in the web server's document root and open `index.php`.
4. Upload the monthly `.xlsx` customer master and then upload the daily IDPEL file.

The application creates its required tables in the configured database. Customer spreadsheets and the live `config.php` are intentionally excluded from Git.
