import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Equation for fitting
def fitting_curve_trunc(x, a, m, k):
    return a * ((x + 1e-6)**(-m)) * np.exp(-k * (x + 1e-6))

def plot_trunc():
    global file_path
    global data

    x_data = data['OnX'][3::].values
    y_data = data['OnY'][3::].values



    # Initial guesses 
    initial_guess = [1.0, 1.0, 1.0]

    # plotdec = input("Truncated plot(1) or non trunca1




    params, covariance = curve_fit(fitting_curve_trunc, x_data, y_data, p0=initial_guess)


    fitted_a, fitted_m, fitted_k = params

    print("Fitted a:", fitted_a)
    print("Fitted m:", fitted_m)
    print("Fitted k:", fitted_k)


    fitted_y = fitting_curve_trunc(x_data, fitted_a, fitted_m, fitted_k)


    # plt.scatter(x_data, y_data, label='Data')
    # plt.plot(x_data, fitted_y, color='red', label='Fitted Curve')
    # plt.xlabel('X')
    # plt.ylabel('Y')
    # plt.title('Curve Fitting')
    # plt.legend()
    # plt.grid(True)
    # plt.show()

    plt.figure(figsize=(8, 6))
    plt.loglog(x_data, y_data, 'bo', label='Scattered data')
    plt.loglog(x_data, fitted_y, 'r-', label='Fitted Curve')
    plt.xlabel('Log(X)')
    plt.ylabel('Log(Y)')
    plt.title('Log-Log Scale Curve Fitting')
    plt.legend()
    plt.grid(True)
    plt.annotate(f'a = {fitted_a:.3f}', xy=(0.1, 0.22), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    plt.annotate(f'm = {fitted_m:.3f}', xy=(0.1, 0.15), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    plt.annotate(f'k = {fitted_k:.3f}', xy=(0.1, 0.08), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    plt.show()

    fitted_data = pd.DataFrame({'x': x_data, 'fitted_y': fitted_y})


    with pd.ExcelWriter('fittedplot.xlsx', engine='xlsxwriter') as writer:
        fitted_data.to_excel(writer, sheet_name='fittedplot', index=False)

    print("Fitted data exported to 'fittedplot.xlsx'")




    # fitted_data = pd.DataFrame({'x': x_data, 'fitted_y': fitted_y})
    # with pd.ExcelWriter('fittedplot_Truncated.xlsx', engine='xlsxwriter') as writer:
    #  fitted_data.to_excel(writer, sheet_name='fittedplot_Truncated', index=False)
    # print("finished")


def fitting_curve_nontrunc(x,a,m):
    return a * ((x + 1e-6)**(-m)) 

def plot_non_trunc():
    global file_path
    global data

    x_data = data['OnX'][3::].values
    y_data = data['OnY'][3::].values



    # Initial guesses 
    initial_guess = [1.0, 1.0]

    # plotdec = input("Truncated plot(1) or non trunca1




    params, covariance = curve_fit(fitting_curve_nontrunc, x_data, y_data, p0=initial_guess)


    fitted_a, fitted_m = params

    print("Fitted a:", fitted_a)
    print("Fitted m:", fitted_m)
    # print("Fitted k:", fitted_k)


    fitted_y = fitting_curve_nontrunc(x_data, fitted_a, fitted_m)


    # plt.scatter(x_data, y_data, label='Data')
    # plt.plot(x_data, fitted_y, color='red', label='Fitted Curve')
    # plt.xlabel('X')
    # plt.ylabel('Y')
    # plt.title('Curve Fitting')
    # plt.legend()
    # plt.grid(True)
    # plt.show()

    plt.figure(figsize=(8, 6))
    plt.loglog(x_data, y_data, 'bo', label='Scattered data')
    plt.loglog(x_data, fitted_y, 'r-', label='Fitted Curve')
    plt.xlabel('Log(X)')
    plt.ylabel('Log(Y)')
    plt.title('Log-Log Scale Curve Fitting')
    plt.legend()
    plt.grid(True)
    plt.annotate(f'a = {fitted_a:.3f}', xy=(0.1, 0.22), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    plt.annotate(f'm = {fitted_m:.3f}', xy=(0.1, 0.15), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    # plt.annotate(f'k = {fitted_k:.3f}', xy=(0.1, 0.08), xycoords='axes fraction', fontsize=12, bbox=dict(boxstyle='round', facecolor='white', edgecolor='black', alpha=0.7))
    plt.show()

    fitted_data = pd.DataFrame({'x': x_data, 'fitted_y': fitted_y})


    with pd.ExcelWriter('fittedplot_NonTruncated.xlsx', engine='xlsxwriter') as writer:
     fitted_data.to_excel(writer, sheet_name='fittedplot_NonTruncated', index=False)



global file_path
global data
root = tk.Tk()
root.withdraw()  

#pop up
file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
data = pd.read_excel(file_path, sheet_name='Sheet1')

# data = pd.read_excel('pdfOn.xlsx', sheet_name='Sheet1')  

# We can select the starting data from which we should read

plotdec = input("Truncated plot(1) or non truncated plot(2):")

dec = int(plotdec)

if dec == 1:
    plot_trunc()

if dec == 2:
    plot_non_trunc()

# fitted_data = pd.DataFrame({'x': x_data, 'fitted_y': fitted_y})


# with pd.ExcelWriter('fittedplot.xlsx', engine='xlsxwriter') as writer:
#     fitted_data.to_excel(writer, sheet_name='fittedplot', index=False)

# print("Fitted data exported to 'fittedplot.xlsx'")
