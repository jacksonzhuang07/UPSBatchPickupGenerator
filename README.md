# 🚀 UPS Pickup Automation Engine

**Transforming Logistics Workflows with Intelligent Automation**

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![API-UPS](https://img.shields.io/badge/API-UPS%20REST-orange.svg)](https://www.ups.com/upsdeveloperkit)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 🌟 Executive Summary
The **UPS Pickup Automation Engine** is a high-performance desktop application designed to eliminate the friction of scheduling high-volume logistics pickups. By leveraging intelligent NLP-based address parsing and direct integration with UPS RESTful APIs, this tool reduces complex manual scheduling tasks from **minutes to seconds**.

---

## 💼 Business Value & Impact

### ⏳ 90% Reduction in Processing Time
Manual entry of a single pickup can take 2-3 minutes. Our **Batch Processing Engine** handles up to 150 pickups in a single click, saving hours of manual labor every week for logistics teams.

### 🎯 Zero-Error Precision
By automating the extraction of addresses and tracking numbers directly from client emails or manifests, the engine eliminates costly human errors such as typos in postal codes or house numbers that lead to missed pickups.

### 📈 Seamless Scalability
Designed for growth, the architecture supports batch operations and handles complex regional logic (US/Canada service codes) automatically, allowing your team to focus on high-value logistics strategy rather than data entry.

### 📋 Enterprise Auditing
Every action is serialized into a robust history with Excel export capabilities, providing immediate transparency for billing, performance tracking, and client reporting.

---

## 🛠️ Technology Stack

| Layer | Technology | Purpose |
| :--- | :--- | :--- |
| **Core** | Python 3.10+ | Robust, scalable application logic. |
| **Parsing** | `usaddress` + Custom Heuristics | NLP-driven extraction of Company, Contact, and Address. |
| **UI** | Tkinter | Lightweight, cross-platform Desktop GUI. |
| **Integrations**| UPS OAuth2 REST APIs | Secure, real-time pickup & shipment management. |
| **Data** | JSON / OpenPyXL | High-speed local history & Professional Excel reporting. |
| **DevOps** | Git / Python-Dotenv | Professional version control and secure credential management. |

---

## 🚀 Key Features

### 🧠 Intelligent Address Parsing
Our custom-built parser uses heuristic logic to intelligently identify:
- **Company Name & Contact Name** from raw, unstructured text.
- **Improved Phone Detection** across global formats.
- **Smart Tracking Detection** for automated 1Z lookup.

### 📦 Automated Label Generation
Integrates with the UPS Shipping API to automatically generate unique return labels when no tracking number is provided, creating a truly "one-click" experience.

### 🌍 Timezone-Aware Scheduling
Automatically calculates local pickup times based on the destination province (NL to BC), ensuring compliance with UPS cutoff windows across multiple time zones.

### 📊 Professional History Engine
- **Searchable Logging**: Filter by PRN, Address, or Company.
- **Live Status Verification**: One-click lookup of real-time UPS status for any scheduled pickup.
- **Batch Export**: Professional-grade Excel reports for auditing and bookkeeping.

---

## 🛠️ Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/UPSPickupAPI.git
   ```
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Configure Environment:**
   Create a `.env` file based on `.env.template` with your UPS API credentials.
4. **Run the App:**
   ```bash
   python main.py
   ```

---

## 👨‍💻 Technical Highlights
- **Asynchronous Execution**: Uses Python threading to ensure the GUI remains responsive during long-running API batch calls.
- **Stateless API Design**: Implements robust OAuth2 token management.
- **Modular Architecture**: Clean separation between the `UPSApiClient`, `AddressParser`, and `GUI` layers for easy maintenance and testing.

---
*Created with focus on efficiency, reliability, and visual excellence.*
