# NPMD Outage Monitoring

A small PHP and MySQL application for monitoring PSA office/project internet
outages. The application stores actual outage events rather than creating
empty records for every day.

## Requirements

- PHP 8+
- MySQL or MariaDB
- PHP PDO MySQL driver

## Run with XAMPP

1. Copy this folder into `C:\xampp\htdocs\`.
2. Start Apache and MySQL from the XAMPP Control Panel.
3. Create a MySQL database named `npmd` in phpMyAdmin.
4. Open `http://localhost/NPMD_OUTAGE_MONITORING/`.

The application creates or updates its tables when it first loads. The default
database connection settings are in `config.php`; adjust them for your local
environment as needed.

## Run with PHP's built-in server

With PHP and MySQL installed and MySQL running, open a terminal in this folder
and run:

```text
php -S 0.0.0.0:8000
```

Then open `http://localhost:8000/`.

## Excel report

Keep `Excel_report_template.xlsx` in the project folder. The Analysis Summary's
Export Excel Report uses its Analysis, Chart Data, and Outage Details sheets.

## Local data files

The populated database dump, SQLite runtime database, and inventory workbook
are excluded from this repository because they may contain operational data.
Back up the database separately using phpMyAdmin or a database dump.

## Hosting note

This is a PHP/MySQL application. GitHub stores the source code, but GitHub Pages
does not execute PHP or provide a MySQL database. To make the application
available on the internet, deploy it to a PHP hosting provider with MySQL.
