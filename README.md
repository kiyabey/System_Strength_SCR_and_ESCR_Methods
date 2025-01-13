# Short Circuit Ratio (SCR) and Equivalent Short Circuit Ratio (ESCR) Calculation Tool

This Python tool calculates the **Short Circuit Ratio (SCR)** and **Equivalent Short Circuit Ratio (ESCR)**, which are commonly used methods for assessing the **system strength** of electrical systems. 
The **SCR** method is typically used for grids dominated by synchronous generators or systems with low penetration rates of inverter-based devices. 
The **ESCR**, on the other hand, extends the SCR calculation by incorporating an **Interaction Factor (IF)** to better account for the interactions between close, inverter-based devices.

## Features

- **Input Format**: Accepts input data in Excel format.
- **SCR and ESCR Calculation**: Calculates both **SCR** and **ESCR** for each bus, with ESCR enhanced by an **Interaction Factor (IF)** to account for interactions in inverter-based grids.
- **Interaction Factor (IF)**: The ESCR method uses an interaction factor to properly account for the interactions between electrically close devices, particularly inverter-based systems.
- **Output**: Saves the calculated SCR and ESCR values in the same Excel file.
- **Mathematical Calculations**: Utilizes mathematical formulas for SCR and ESCR, with an added factor for system interactions in the ESCR calculation.

### **Key Formulas**

- **SCR**: The Short Circuit Ratio (SCR) at node \(i\) is calculated as:  
  \[
  SCR_i = \frac{\text{Subtransient Short-Circuit Power at node } i}{\text{Nominal Power of the Inverters at node } i}
  \]
  This formula represents the ratio of the subtransient short-circuit power at node \(i\) to the nominal power of the inverters at the same node.

- **ESCR**: The Equivalent Short Circuit Ratio (ESCR) at node \(i\) is calculated as:  
  \[
  ESCR_i = \frac{\text{Subtransient Short-Circuit Power at node } i}{\text{Nominal Power of the Inverter at node } i + \sum_{i \neq j} (\text{Interaction Factor between nodes } i \text{ and } j \times \text{Nominal Power of the Inverter at node } j )}
  \]
  The **ESCR** incorporates the interaction factor (IF) between node \(i\) and other nodes \(j\), accounting for the interaction effects between inverter-based devices at different nodes.

- **Interaction Factor (IF)**: The interaction factor is a system-dependent value that accounts for the electrical interaction between inverter-based devices at different nodes.
- The interaction factor modifies the ESCR calculation to represent these interactions more accurately. It can be seen as the voltage sensitivity among the buses of the system.

## Requirements

- **Python Version**: Python 3.8 or later.
- **Dependencies**:
  - `openpyxl`: For reading and writing Excel files.
  - `math`: For mathematical calculations (e.g., square root for ESCR calculation).

Install the dependencies using:
```bash
pip install openpyxl

#### **Contact**

If you have any questions, feel free to reach out:
- **Linkedin**: www.linkedin.com/in/bartu-badem-4a726b283

