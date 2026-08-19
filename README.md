# Benchmarks Web Server

Web server hosting the landing page for performance benchmark tracking and metrics display.

## Ticket Reference
* **Jira Ticket:** [PERF-527] Create a landing page on the benchmarks web server

## Overview
This service serves as the central entry point for benchmark reports and performance metrics across the system.

## Setup and Installation

### Prerequisites
Make sure you have **Python 3.8+** and **Git** installed on your system.

### Clone the Repository
git clone https://github.com/mariadb-rahulraj/benchmark-landing-page.git

### Install all required packages at once
pip install -r requirements.txt

### Run the Streamlit application
python3 -m streamlit run landing_page.py
