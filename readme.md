# EntraSecure – A Guest User Visibility & Governance Dashboard using Python and Microsoft Graph

**EntraSecure** is an automated identity governance tool designed to provide real-time visibility into guest user accounts in Microsoft Entra ID (formerly Azure Active Directory). It leverages Microsoft Graph API to retrieve guest user data, visualizes that data in Excel Online, and delivers interactive reports directly to Microsoft Teams.

Built with Python and structured for automation, this project emphasizes secure access, data-driven auditing, and seamless integration into Microsoft 365 environments.

## Identity & Security Focus

- **Graph API + Secure Auth**  
  Utilizes OAuth 2.0 client credentials flow to authenticate with Microsoft Graph and interact with Entra ID securely.

- **Guest Access Visibility**  
  Surfaces key metadata about external users for lifecycle management, risk analysis, and policy alignment.

- **Excel-Based Dashboarding**  
  Writes guest data to a live OneDrive-hosted Excel file, generating a dynamic chart for ongoing analysis.

- **Automated Reporting to Microsoft Teams**  
  Sends a snapshot of the current guest user chart (base64 image) directly to a Teams channel for visibility by IT, GRC, or SOC teams.

## Project Features

- **Entra ID Guest User Querying**  
  Scans the directory for all guest accounts (`userType = Guest`) using Microsoft Graph.

- **Excel Table Population**  
  Appends user info (name, email, created date) into a formatted Excel Online table.

- **Chart Creation and Export**  
  Automatically generates a clustered column chart based on user account age and distribution.

- **Teams Integration**  
  Posts a rich HTML message containing the generated chart image to a selected Teams channel.

## Architecture Overview

### Python Automation

- **Platform**: Python 3.10+  
- **Graph SDK**: Raw REST with `requests`  
- **Security**:
  - Client credentials flow (no user interaction)
  - No local password storage
  - Principle of least privilege (read-only for Entra + write to Excel/Teams only)

### Microsoft 365 Integration

- **Data Source**: Microsoft Entra ID guest user directory  
- **Report Destination**: Excel Online (OneDrive for Business)  
- **Delivery Channel**: Microsoft Teams via Graph message API

## Setup Instructions

### Prerequisites

- Azure AD tenant with:
  - OneDrive for Business and Teams access
  - A registered App Registration in Azure Active Directory
- API permissions (Application type):
  - `User.Read.All`
  - `Group.Read.All`
  - `Files.ReadWrite.All`
  - `ChannelMessage.Send`
- Admin consent granted

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/GrantPierce94/entrasecure.git
   cd entrasecure
   pip install -r requirements.txt
