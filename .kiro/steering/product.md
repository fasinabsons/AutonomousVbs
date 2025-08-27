# MoonFlower WiFi Automation System

## Product Overview

MoonFlower is a comprehensive WiFi automation system designed for 365-day continuous operation. The system automates the complete workflow of WiFi data collection, processing, and reporting for network management.

## Core Functionality

- **Automated CSV Downloads**: Scheduled downloads from Ruckus Wireless Controller (4 time slots daily)
- **Excel Processing**: Converts CSV data to VBS-compatible Excel format (.xls) with proper header mapping
- **VBS Automation**: 4-phase automation system (Login, Navigation, Upload, Report Generation)
- **Windows Service**: Background service operation with auto-startup and health monitoring
- **Email Notifications**: Automated daily reports and error notifications
- **PC Restart Management**: Intelligent system restart scheduling based on health metrics

## Key Features

- **Reliability**: Multiple startup mechanisms, health monitoring, auto-recovery
- **Scheduling**: Time-based task execution with dependency management
- **Error Handling**: Comprehensive retry logic and graceful degradation
- **Monitoring**: Real-time health checks and performance tracking
- **Maintenance**: Automatic log rotation, file cleanup, and system optimization

## Target Environment

- **Platform**: Windows systems with Chrome browser
- **Operation Mode**: Unattended background service
- **Data Sources**: Ruckus Wireless Controller web interface
- **Output**: Excel files, PDF reports, email notifications