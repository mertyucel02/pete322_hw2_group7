import xlwings as xw
import re
import math
from tabulate import tabulate
import numpy as np
from matplotlib import pyplot as plt

bookName = r'C:\Users\MERT\Desktop\pete322_CasingDesignProject\Casing-Data-Sheet.xlsx'
sheetName = 'Casing Data'

wb = xw.Book(bookName)
sht = wb.sheets[sheetName]
num_list = []
exit_code = 1
csg_size = float(input("Enter Casing Size (in): "))
while exit_code != 0:
    csg_grade = str(input("Enter Casing Grade: "))

    match = re.search(r'\d+', csg_grade)
    if match:
        csg_grade_number = int(match.group())
        min_yield_strength = csg_grade_number * 1000
        print("Minimum Yield Strength: {} psi".format(min_yield_strength))
    else:
        print("No integer part found in the casing grade.")
        csg_grade_number = None
        min_yield_strength = None

    if csg_grade_number is not None:
        grade_cell = sht.api.UsedRange.Find(csg_grade, MatchCase=False)

        row_numbers = []

        if grade_cell is not None:

            if grade_cell.Value.lower() == csg_grade.lower():
                row_numbers.append(grade_cell.Row)

            initial_address = grade_cell.Address

            while True:
                grade_cell = sht.api.UsedRange.FindNext(grade_cell)
                if grade_cell is None or grade_cell.Address == initial_address:
                    break

                if grade_cell.Value.lower() == csg_grade.lower():
                    row_numbers.append(grade_cell.Row)

            if row_numbers:
                print("Rows containing '{}' grade:".format(csg_grade), row_numbers)

                row_numbers_filtered = []
                for row_num in row_numbers:
                    # Get the casing size from column A
                    casing_size = sht.range(f"A{row_num}").value
                    if casing_size == csg_size:
                        row_numbers_filtered.append(row_num)

                if row_numbers_filtered:
                    print("Row numbers corresponding to casing size {} in.: {}".format(csg_size, row_numbers_filtered))

                    csg_id_list = []
                    nom_w_values = []
                    for row_num in row_numbers_filtered:
                        csg_id = sht.range(f"N{row_num}").value
                        csg_id_list.append(csg_id)
                        # Get the value from column B
                        nom_w = sht.range(f"B{row_num}").value
                        nom_w_values.append(nom_w)

                    print("Inner diameters for the filtered rows: {} in.".format(csg_id_list))
                else:
                    raise ValueError("No rows found for casing size {} in.".format(csg_size))
            else:
                raise ValueError("No rows containing '{}' grade found.".format(csg_grade))
        else:
            raise ValueError("Casing grade '{}' not found.".format(csg_grade))
    else:
        raise ValueError("Invalid casing grade '{}' entered.".format(csg_grade))

    c_0 = 2.8762
    c_1 = .10679 * math.pow(10, -10)
    c_2 = .021302 * math.pow(10, -10)
    c_3 = -.53132 * math.pow(10, -16)
    c_4 = .026233
    c_5 = .50609 * math.pow(10, -6)
    c_6 = -465.93
    c_7 = .030867
    c_8 = -.10483 * math.pow(10, -7)
    c_9 = .36989 * math.pow(10, -13)
    c_10 = 46.95 * math.pow(10, 6)

    F_1 = c_0 + c_1 * min_yield_strength + c_2 * math.pow(min_yield_strength, 2) + c_3 * math.pow(min_yield_strength, 3)
    F_2 = c_4 + c_5 * min_yield_strength
    F_3 = c_6 + c_7 * min_yield_strength + c_8 * math.pow(min_yield_strength, 2) + c_9 * math.pow(min_yield_strength, 3)
    R_F = F_2 / F_1
    F_4 = (c_10 * math.pow(3 * R_F / (2 + R_F), 3)) / (min_yield_strength * (3 * R_F / (2 + R_F) - R_F) * math.pow(1 - 3 * R_F / (2 + R_F), 2))
    F_5 = F_4 * R_F

    d_n_over_t_yield = (math.sqrt(math.pow(F_1 - 2, 2) + 8 * (F_2 + F_3 / min_yield_strength)) + (F_1 - 2)) / (2 * (F_2 + F_3 / min_yield_strength))

    d_n_over_t_plastic = (2 + F_2 / F_1) / (3 * F_2 * F_1)

    d_n_over_t_transition = (min_yield_strength * (F_1 - F_4)) / (F_3 + min_yield_strength * (F_2 - F_5))

    d_over_t_list = []
    p_cr_yield_list = []
    p_cr_plastic_list = []
    p_cr_transition_list = []
    p_cr_elastic_list = []

    for i in range(0, len(row_numbers_filtered)):
        d_over_t = 2 * csg_size / (csg_size - csg_id_list[i])
        d_over_t_list.append(d_over_t)
        if d_n_over_t_yield >= d_over_t:
            p_cr_yield = 2 * min_yield_strength * ((d_over_t - 1) / math.pow(d_over_t, 2))
        else:
            p_cr_yield = str("-")
        p_cr_yield_list.append(p_cr_yield)

        if d_n_over_t_plastic >= d_over_t:
            p_cr_plastic = min_yield_strength * (F_1 / d_over_t - F_2) - F_3
        else:
            p_cr_plastic = str("-")
        p_cr_plastic_list.append(p_cr_plastic)

        if d_n_over_t_transition <= d_over_t:
            p_cr_transition = min_yield_strength * ((F_4 / d_over_t) - F_5)
        else:
            p_cr_transition = str("-")
        p_cr_transition_list.append(p_cr_transition)

        p_cr_elastic = (46.95 * math.pow(10, 6)) / (d_over_t * math.pow(d_over_t - 1, 2))
        p_cr_elastic_list.append(p_cr_elastic)

    my_data = []
    for i in range(len(row_numbers_filtered)):
        my_data.append([i + 1, csg_size, csg_grade, nom_w_values[i], min_yield_strength, csg_id_list[i], d_over_t_list[i], p_cr_yield_list[i], p_cr_plastic_list[i], p_cr_transition_list[i], p_cr_elastic_list[i]])
    head = ["#", "OD, in", "Casing Grade",  "Nominal Weight, #/ft", "Minimum Yield Strength, psi", "ID, in", "d/t", "Yield Strength Collapse, psi", "Plastic Collapse, psi",
            "Transition Collapse, psi", "Elastic Collapse, psi"]

    print(tabulate(my_data, headers=head, tablefmt="grid"))

    csg_p_b_list = []
    csg_p_c_list = []

    for i in range(len(row_numbers_filtered)):
        csg_p_b = 0.875 * 2 * min_yield_strength * ((csg_size - csg_id_list[i]) / 2) / csg_size
        csg_p_b_list.append(csg_p_b)
        if isinstance(p_cr_yield_list[i], float):
            csg_p_c_list.append(p_cr_yield_list[i])
        elif isinstance(p_cr_plastic_list[i], float):
            csg_p_c_list.append(p_cr_plastic_list[i])
        elif isinstance(p_cr_transition_list[i], float):
            csg_p_c_list.append(p_cr_transition_list[i])
        else:
            csg_p_c_list.append(p_cr_elastic_list[i])

    my_data_2 = []
    for i in range(len(row_numbers_filtered)):
        my_data_2.append([i + 1, csg_size, csg_grade, nom_w_values[i], min_yield_strength, csg_id_list[i], csg_p_b_list[i], csg_p_c_list[i]])

    head_2 = ["#", "OD, in", "Casing Grade",  "Nominal Weight, #/ft", "Minimum Yield Strength, psi", "ID, in", "Burst Pressure, psi", "Collapse Pressure, psi"]

    print(tabulate(my_data_2, headers=head_2, tablefmt="grid"))

    num = int(input("Enter Selected Casing Number: "))
    num_list.append(my_data_2[num - 1][1:])

    to_exit = int(input("Enter 0 to exit, otherwise enter anything integer excepted 0: "))
    if to_exit == 0:
        break
    else:
        continue

head_3 = ["OD, in", "Casing Grade",  "Nominal Weight, #/ft", "Minimum Yield Strength, psi", "ID, in", "Burst Pressure, psi", "Collapse Pressure, psi"]
print(tabulate(num_list, headers=head_3, tablefmt="grid"))
design_depth = int(input("Enter Design Depth (ft): "))
mud_wt = float(input("Enter Mud Weight (ppg): "))
PG = float(input("Enter Pore Pressure Gradient (psi/ft): "))
TF_max = float(input("Enter Maximum Tensile Force (lbf): "))
print("------SAFETY FACTORS------")
SF_yield = float(input("Tension and Joint Strength: "))
SF_collapse = float(input("Collapse: "))
SF_burst = float(input("Burst: "))
tension = TF_max * SF_yield
p_b_req = PG * SF_burst * design_depth
p_c_req = .052 * mud_wt * design_depth * SF_collapse
p_b_test_csg = 0
p_c_test_csg = 0
i_of_p_test = 0
for i in range(len(num_list)):
    if i == 0:
        if num_list[i][5] >= p_b_req and num_list[i][6] >= p_c_req:
            p_b_test_csg = num_list[i][5]
            p_c_test_csg = num_list[i][6]
            i_of_p_test = i
        else:
            p_b_test_csg = 0
            p_c_test_csg = 0
    elif 0 < i < (len(num_list) - 1):
        if p_b_test_csg == 0 and p_c_test_csg == 0:
            if num_list[i][5] >= p_b_req and num_list[i][6] >= p_c_req:
                p_b_test_csg = num_list[i][5]
                p_c_test_csg = num_list[i][6]
                i_of_p_test = i
            else:
                p_b_test_csg = 0
                p_c_test_csg = 0
        else:
            if p_b_test_csg >= num_list[i][5] >= p_b_req and p_c_test_csg >= num_list[i][6] >= p_c_req:
                p_b_test_csg = num_list[i][5]
                p_c_test_csg = num_list[i][6]
                i_of_p_test = i
            else:
                continue
    elif i == len(num_list) - 1:
        if p_b_test_csg == 0 and p_c_test_csg == 0:
            if num_list[i][5] >= p_b_req and num_list[i][6] >= p_c_req:
                p_b_test_csg = num_list[i][5]
                p_c_test_csg = num_list[i][6]
                i_of_p_test = i
            else:
                raise ValueError("Required Burst or Collapse Pressures is NOT SUITABLE.")
        elif p_b_test_csg >= num_list[i][5] >= p_b_req and p_c_test_csg >= num_list[i][6] >= p_c_req:
            p_b_test_csg = num_list[i][5]
            p_c_test_csg = num_list[i][6]
            i_of_p_test = i
        else:
            break
csg_grade_test_list = []
p_b_test_list = []
p_c_test_list = []
nom_w_test_list = []
ID_test_list = []
min_yield_strength_test_list = []

for i in range(len(num_list)):
    if num_list[i][2] <= num_list[i_of_p_test][2]:
        if num_list[i][5] <= p_b_test_csg:
            csg_grade_test_list.append(num_list[i][1])
            p_b_test_list.append(num_list[i][5])
            p_c_test_list.append(num_list[i][6])
            nom_w_test_list.append(num_list[i][2])
            ID_test_list.append(num_list[i][4])
            min_yield_strength_test_list.append(num_list[i][3])
        else:
            continue
    else:
        continue


combined_list = list(zip(csg_grade_test_list, nom_w_test_list, p_b_test_list, p_c_test_list, ID_test_list, min_yield_strength_test_list))
sorted_combined_list = sorted(combined_list, key=lambda x: x[1], reverse=True)
sorted_csg_grade_test_list, sorted_nom_w_test_list, sorted_p_b_test_list, sorted_p_c_test_list, sorted_ID_test_list, sorted_min_yield_strength_test_list = map(list, zip(*sorted_combined_list))
res_list = []
h_2 = design_depth
for i in range(len(sorted_nom_w_test_list)):
    p_c_1 = sorted_p_c_test_list[i] / SF_collapse
    h_1 = p_c_1 / (.052 * mud_wt)
    W_1 = sorted_nom_w_test_list[i] * (h_2 - h_1)
    S_1 = W_1 / (math.pi * (math.pow(csg_size, 2) - math.pow(sorted_ID_test_list[i], 2)) / 4)
    Y_PA_1 = sorted_min_yield_strength_test_list[i] * (math.sqrt(1 - .75 * math.pow(S_1 / sorted_min_yield_strength_test_list[i], 2)) - .5 * (S_1 / sorted_min_yield_strength_test_list[i]))
    c_0 = 2.8762
    c_1 = .10679 * math.pow(10, -10)
    c_2 = .021302 * math.pow(10, -10)
    c_3 = -.53132 * math.pow(10, -16)
    c_4 = .026233
    c_5 = .50609 * math.pow(10, -6)
    c_6 = -465.93
    c_7 = .030867
    c_8 = -.10483 * math.pow(10, -7)
    c_9 = .36989 * math.pow(10, -13)
    c_10 = 46.95 * math.pow(10, 6)

    F_1 = c_0 + c_1 * Y_PA_1 + c_2 * math.pow(Y_PA_1, 2) + c_3 * math.pow(Y_PA_1, 3)
    F_2 = c_4 + c_5 * Y_PA_1
    F_3 = c_6 + c_7 * Y_PA_1 + c_8 * math.pow(Y_PA_1, 2) + c_9 * math.pow(Y_PA_1, 3)
    R_F = F_2 / F_1
    F_4 = (c_10 * math.pow(3 * R_F / (2 + R_F), 3)) / (Y_PA_1 * (3 * R_F / (2 + R_F) - R_F) * math.pow(1 - 3 * R_F / (2 + R_F), 2))
    F_5 = F_4 * R_F

    d_n_over_t_yield = (math.sqrt(math.pow(F_1 - 2, 2) + 8 * (F_2 + F_3 / Y_PA_1)) + (F_1 - 2)) / (2 * (F_2 + F_3 / Y_PA_1))

    d_n_over_t_plastic = (2 + F_2 / F_1) / (3 * F_2 * F_1)

    d_n_over_t_transition = (Y_PA_1 * (F_1 - F_4)) / (F_3 + Y_PA_1 * (F_2 - F_5))
    d_over_t = 2 * csg_size / (csg_size - sorted_ID_test_list[i])
    if d_n_over_t_yield >= d_over_t:
        p_cr_yield = 2 * Y_PA_1 * ((d_over_t - 1) / math.pow(d_over_t, 2))
    else:
        p_cr_yield = str("-")

    if d_n_over_t_plastic >= d_over_t:
        p_cr_plastic = Y_PA_1 * (F_1 / d_over_t - F_2) - F_3
    else:
        p_cr_plastic = str("-")

    if d_n_over_t_transition <= d_over_t:
        p_cr_transition = Y_PA_1 * ((F_4 / d_over_t) - F_5)
    else:
        p_cr_transition = str("-")

    p_cr_elastic = (46.95 * math.pow(10, 6)) / (d_over_t * math.pow(d_over_t - 1, 2))

    if isinstance(p_cr_yield, float):
        p_cr_1 = p_cr_yield
    elif isinstance(p_cr_plastic, float):
        p_cr_1 = p_cr_plastic
    elif isinstance(p_cr_transition, float):
        p_cr_1 = p_cr_transition
    else:
        p_cr_1 = p_cr_elastic

    p_c_2 = p_cr_1 / SF_collapse
    h_2 = p_c_2 / (.052 * mud_wt)
    Y_PA_previous = Y_PA_1
    while np.abs(h_2 - h_1) > 30:
        p_c_1 = p_c_2
        h_1 = p_c_1 / (.052 * mud_wt)
        W_1 = sorted_nom_w_test_list[i] * (h_2 - h_1)
        S_1 = W_1 / (math.pi * (math.pow(csg_size, 2) - math.pow(sorted_ID_test_list[i], 2)) / 4)
        Y_PA_1 = Y_PA_previous * (math.sqrt(1 - .75 * math.pow(S_1 / sorted_min_yield_strength_test_list[i], 2)) - .5 * (S_1 / sorted_min_yield_strength_test_list[i]))
        c_0 = 2.8762
        c_1 = .10679 * math.pow(10, -10)
        c_2 = .021302 * math.pow(10, -10)
        c_3 = -.53132 * math.pow(10, -16)
        c_4 = .026233
        c_5 = .50609 * math.pow(10, -6)
        c_6 = -465.93
        c_7 = .030867
        c_8 = -.10483 * math.pow(10, -7)
        c_9 = .36989 * math.pow(10, -13)
        c_10 = 46.95 * math.pow(10, 6)

        F_1 = c_0 + c_1 * Y_PA_1 + c_2 * math.pow(Y_PA_1, 2) + c_3 * math.pow(Y_PA_1, 3)
        F_2 = c_4 + c_5 * Y_PA_1
        F_3 = c_6 + c_7 * Y_PA_1 + c_8 * math.pow(Y_PA_1, 2) + c_9 * math.pow(Y_PA_1, 3)
        R_F = F_2 / F_1
        F_4 = (c_10 * math.pow(3 * R_F / (2 + R_F), 3)) / (Y_PA_1 * (3 * R_F / (2 + R_F) - R_F) * math.pow(1 - 3 * R_F / (2 + R_F), 2))
        F_5 = F_4 * R_F

        d_n_over_t_yield = (math.sqrt(math.pow(F_1 - 2, 2) + 8 * (F_2 + F_3 / Y_PA_1)) + (F_1 - 2)) / (2 * (F_2 + F_3 / Y_PA_1))

        d_n_over_t_plastic = (2 + F_2 / F_1) / (3 * F_2 * F_1)

        d_n_over_t_transition = (Y_PA_1 * (F_1 - F_4)) / (F_3 + Y_PA_1 * (F_2 - F_5))
        d_over_t = 2 * csg_size / (csg_size - sorted_ID_test_list[i])
        if d_n_over_t_yield >= d_over_t:
            p_cr_yield = 2 * Y_PA_1 * ((d_over_t - 1) / math.pow(d_over_t, 2))
        else:
            p_cr_yield = str("-")

        if d_n_over_t_plastic >= d_over_t:
            p_cr_plastic = Y_PA_1 * (F_1 / d_over_t - F_2) - F_3
        else:
            p_cr_plastic = str("-")

        if d_n_over_t_transition <= d_over_t:
            p_cr_transition = Y_PA_1 * ((F_4 / d_over_t) - F_5)
        else:
            p_cr_transition = str("-")

        p_cr_elastic = (46.95 * math.pow(10, 6)) / (d_over_t * math.pow(d_over_t - 1, 2))

        if isinstance(p_cr_yield, float):
            p_cr_1 = p_cr_yield
        elif isinstance(p_cr_plastic, float):
            p_cr_1 = p_cr_plastic
        elif isinstance(p_cr_transition, float):
            p_cr_1 = p_cr_transition
        else:
            p_cr_1 = p_cr_elastic

        p_c_2 = p_cr_1 / SF_collapse
        h_2 = p_c_2 / (.052 * mud_wt)
        Y_PA_previous = Y_PA_1
    res_list.append(h_2)


res_list.append(0)
res_list[0] = design_depth
final_data = []
for i in range(len(res_list)-1):
    final_data.append([i+1, csg_size, sorted_csg_grade_test_list[i], sorted_nom_w_test_list[i], res_list[i], res_list[i+1], res_list[i] - res_list[i+1]])

head_final = ["#", "OD, in", "Grade", "Nominal Weight, #/ft", "Start Depth, ft", "Final Depth, ft", "Displacement, ft"]
print(tabulate(final_data, headers=head_final, tablefmt="grid"))

tension_check_tot = 0
tension_check_list = []
for i in range(len(res_list) - 1):
    tension_check_list.append(SF_yield * (res_list[i] - res_list[i+1]) * sorted_nom_w_test_list[i])
    tension_check_tot += (SF_yield * (res_list[i] - res_list[i+1]) * sorted_nom_w_test_list[i])

print("With a design factor of " + str(SF_yield) + " for tension, a pipe strength of " + str(tension_check_tot) + " lbf is required.")

list_x = []
list_y = []
for i in range(len(res_list) - 1):
    list_x.append([csg_size, csg_size])
    list_y.append([res_list[i], res_list[i+1]])

for i in range(len(list_x)):
    plt.plot(list_x[i], list_y[i], label = (str(csg_size) + " in OD " + str(sorted_csg_grade_test_list[i]) + " " + str(sorted_nom_w_test_list[i]) + " #/ft"))

plt.xlabel("Casing OD ($in$)")
plt.ylabel("Depth ($ft$)")
plt.title("The Plot of Casing Design")
plt.legend()
plt.ylim(bottom = 0)
plt.gca().invert_yaxis()
plt.grid(True)
plt.show()
