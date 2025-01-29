# 📌 Temperature Calibration Tool

## 🔥 Overview

The **Temperature Calibration Tool** automates the calibration process for temperature controllers, reducing calibration time by **50%** compared to manual methods. It fine-tunes the **offset** and **linearity** parameters to minimize the difference between the simulator's actual temperature and the controller's displayed temperature. The tool also provides graphical analysis and automated documentation for customers.

## 🚀 Features

- **Automated Calibration:** Tunes **offset (0-1)** and **linearity (0-1)** parameters for precise temperature matching.
- **Graphical Representation:** Generates graphs to visualize temperature differences at various parameter values.
- **Real-time Notifications:** Alerts when calibration is complete and optimal parameters are found.
- **Excel Report Generation:** Automatically creates datasheets detailing the best calibration results for customers.
- **User-friendly Interface:** Built using **Windows Forms Application** for an intuitive experience.

## 🛠️ Technology Stack

- **Programming Language:** C#
- **Framework:** .NET Framework
- **IDE:** Visual Studio
- **UI:** Windows Forms Application
- **Data Storage:** Excel for report generation

## 📖 Installation Guide

### Prerequisites
- **Windows OS** (Recommended: Windows 10 or later)
- **.NET Framework** (Installed with Visual Studio)
- **Microsoft Excel** (For report generation)

### Steps to Install & Run

1. **Clone the Repository:**
   ```sh
   https://github.com/abbhishek4444/Temperature_Calibration_tool.git
   ```
2. **Navigate to the Project Directory:**
   ```sh
   cd Temperature_Calibration_tool
   ```
3. **Open the Solution in Visual Studio:**
   - Launch **Visual Studio** and open the `.sln` file.
4. **Build & Run the Application:**
   - Click on **Start (F5)** in Visual Studio.

## 📷 Screenshots (Optional)
Include screenshots of the calibration process, graphs, and report generation.
![Alt Text](Temperature_Calibration_tool/Calibrationtool.png)

## 📜 How It Works

1. **Set Simulator Temperature:** Input a target temperature (e.g., 25°C).
2. **Read Controller Output:** The controller displays a different temperature (e.g., 23.5°C).
3. **Automatic Parameter Tuning:** The tool iterates through **offset** and **linearity** values (0 to 1 in steps of 0.1) to minimize the difference.
4. **Graph Generation:** Plots temperature variations against parameter values to identify optimal settings.
5. **Calibration Notification:** Displays "Calibration is done. Best parameters found!"
6. **Generate Report:** Exports an Excel file with calibration results for customers.

## ✅ Contribution Guidelines

- Fork the repository.
- Create a new branch for your feature.
- Commit your changes with clear messages.
- Create a pull request for review.

## 📄 License

This project is licensed under [Your Chosen License, e.g., MIT, Apache 2.0].

## 📞 Contact

For any queries or suggestions, feel free to contact:

- **Your Name:** abbhishek4444@gmail.com
- **GitHub:** https://github.com/abbhishek4444

---

**🚀 Happy Calibrating!**

