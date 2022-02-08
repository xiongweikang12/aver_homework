# TODO 读取excel的学分绩点 ，添加科目，预测更新绩点,可视化，web
import tkinter
import tkinter as tk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import time
from PIL import Image,ImageTk
import re

plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示正文标签
plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

filename = "XSCJ_REPORT (2).xls"
sheet_sor = pd.read_excel(filename)[4:-3]
re_index = list(sheet_sor.loc[4].values)
sheet_sor.columns = re_index
sheet_sor = sheet_sor.drop(4, axis=0)
print(sheet_sor)
study_term = list(sheet_sor.loc[:, '学年学期'].values)
learning_cou = list(sheet_sor.loc[:, "课程名称"].values)
study_power = list(sheet_sor.loc[:, "学分"].values)
grade = list(sheet_sor.loc[:, "成绩"].values)
demo=0
for i in range(len(learning_cou)):
    try:
        grade[i] = int(grade[i])
    except ValueError:
        # def grade_fail_waring():
        if grade[i]=='不及格':
            # print("{0}，第{1}个，为{2}".format(learning_cou[i], i, grade[i]))
            demo=i
            grade[i] = 48
        elif grade[i]=='良好':
            grade[i]=88
        elif grade[i]=='中等':
            grade[i]=78
        pass
    else:
        grade[i] = int(grade[i])
print(demo)

container_dict = []
for j in range(len(learning_cou)):
    cour_digit = {learning_cou[j]: (grade[j], study_power[j])}
    container_dict.append(cour_digit)

# def grade_point(list1, list2, list3):  # list1(名称），list2(学分），list3（成绩）
#     arr_1 = np.array(list2)
#     arr_2 = np.array(list3)
#     arr_2 = (arr_2 - 48) / 10
#     arr_3 = arr_1 * arr_2
#     sum_arr_3 = sum(arr_3)
#     sum_arr_1 = sum(arr_1)
#     GradePoint = sum_arr_3 / sum_arr_1
#
#     return GradePoint

print(container_dict)
print(study_term)
container_all_dict = []
for k in range(len(study_term)):
    all_dict = {study_term[k]: container_dict[k]}
    container_all_dict.append(all_dict)



def point_term(d_dict):  # 返回各学期的学习情况 {xx:(grade,credit)}
    # d_dict为总的{term:{xx:(grade,credit)}} container_all_dict
    # 返回的为container_dict 的各学期部分，为子集
    # 用到了列表推导式
    global choice
    choice=input('学期;')

    if choice == "2021-2022 第一学期":
        term_3 = [p3[choice] for p3 in d_dict if p3.get('2021-2022 第一学期')]
        return term_3
    elif choice == "2020-2021 第二学期":
        term_2 = [p2[choice] for p2 in d_dict if p2.get('2020-2021 第二学期')]
        return term_2
    elif choice == "2020-2021 第一学期":
        term_1 = [p1[choice] for p1 in d_dict if p1.get('2020-2021 第一学期')]
        return term_1
    else:
        pass

def grade_point(al_dict):
    container_list_grade = []
    container_list_credit = []

    # 将列表字典（储存信息）转化为能计算
    for k_2 in al_dict:
        for k_1 in k_2.items():
            container_list_credit.append(k_1[1][1])
            container_list_grade.append(k_1[1][0])
    # for k_3 in range(len(learning_cou)):
    #     container_list_3.append((learning_cou[k_3], container_list_grade[k_3], container_list_credit[k_3]))

    print(container_list_credit)
    print(container_list_grade)

    return container_list_grade, container_list_credit


def con_point(list_g, list_c):  # container_list_grade, container_list_credit
    # 通过输入其学分，成绩，将列表转化为数组计算得到平均绩点
    arr_1 = np.array(list_g)
    arr_2 = np.array(list_c)
    arr_1 = (arr_1 - 48) / 10
    arr_3 = arr_1 * arr_2
    sum_arr_3 = sum(arr_3)
    sum_arr_2 = sum(arr_2)
    GradePoint = sum_arr_3 / sum_arr_2

    return GradePoint


color_lab = ['r', 'b', 'y', 'g']
study_term_set = list(set(study_term))
replace2_ = study_term_set[1]
replace1_ = study_term_set[0]  # 2021第二学期
study_term_set[0] = study_term_set[2]
study_term_set[1] = replace2_
study_term_set[2] = replace1_



def show_content():
    for ki in range(3):
        kr_container = []
        term = point_term(container_all_dict)  # 各学期信息
        c_g, c_c = grade_point(term)  # 通过学期信息获取待计算的数据（用于计算平均绩点）
        kr = con_point(c_g, c_c)  #
        kr_container.append(kr)
        # plt.bar(range(len(c_g)), c_g)
        try:
            plt.subplot(2, 2, ki + 1)
        except Exception:
            pass
        else:
            plt.grid(True)
            plt.axis([0, 10, 0, 100])
            plt.title(study_term_set[ki] + '的学习情况:' + str(kr)[0:4], fontsize=10)
            plt.ylabel('成绩', color='r', fontsize=10)
            plt.plot(range(len(c_g)), c_g, color_lab[ki], linewidth='5.0', alpha=0.5)

    plt.subplot(2, 2, 4)
    plt.grid(True)
    plt.axis([0, 30, 0, 100])
    plt.plot(grade, color_lab[3], linewidth='5.0', alpha=0.5)

    plt.show()


#TODO 创建一个视图窗口，含有文本，按钮控件
root = tk.Tk()
frm_1=tk.Frame(root)
frm_1.pack()
root.title('学年学期')
root.geometry('200x200')
frm_2=tk.Frame(frm_1)
frm_2.pack()
e_1 = tk.Entry(frm_2, show=None)  # 姓名
e_1.pack()

#entry输入名字，b_oppoist的按钮，查询挂科项目，text文本显示
def is_fail():
    var = e_1.get()
    t.insert('end', var)
    var_content = str(learning_cou[demo]) +' '+ str(demo) + ' '+str(grade[demo])
    t.insert('end', var_content)
    # t.insert('end',var_content)
#b_show_graph 按钮 is_show图像分析各个学期学情，学分绩点
def is_show():
    show_content()

frm_3=tk.Frame(frm_1)
frm_3.pack()
frm_button_1=tk.Frame(frm_3)
frm_button_1.pack(side='right')
b_oppoist = tk.Button(frm_button_1, text='学习情况', width=10, height=1, bg='red', command=is_fail)
b_oppoist.pack()
frm_label=tk.Frame(frm_3)
frm_label.pack(side='left')
l_button_1=tk.Label(frm_label,text='学生挂科情况')
l_button_1.pack()
frm_4=tk.Frame(frm_1)
frm_4.pack()
frm_button_2=tk.Frame(frm_4)
frm_button_2.pack(side='right')
frm_label_2=tk.Frame(frm_4)
frm_label_2.pack(side='left')
l_button_2=tk.Label(frm_label_2,text='绩点成绩可视化')
l_button_2.pack()
b_show_graph=tk.Button(frm_button_2,text='学情分析',width=10,height=1,bg='blue',command=is_show)
b_show_graph.pack()
frm_5=tk.Frame(frm_1)
frm_5.pack()
frm_button_3=tk.Frame(frm_5)
frm_button_3.pack(side='right')
frm_label_3=tk.Frame(frm_5)
frm_label_3.pack(side='left')
l_button_3=tk.Label(frm_label_3,text='退出主程序')
l_button_3.pack()
b_quit_all=tk.Button(frm_button_3,text='退出',width=10,height=1,bg='green',command=quit) #退出按钮
b_quit_all.pack()
t=tk.Text(frm_1,height=2)
t.pack()


#TODO 增添框架：listbox各各学期的学习情况(直方图)
frm_11=tk.Frame(frm_1)
frm_11.pack()
l_1=tk.Label(frm_11,text='listbox各各学期的学习情况',bg='yellow')
l_1.pack()
var_l_2=tk.StringVar()
l_2=tk.Label(frm_11,textvariable=var_l_2 ,bg='yellow',width=40,height=1)
l_2.pack()
var_item=tk.StringVar()
img_open=Image.open('动漫.jpg').resize((200,200))
img_jpg=ImageTk.PhotoImage(img_open)
l_3=tk.Label(frm_11,image=img_jpg)
l_3.pack()
var_item.set(('2020-2021 第一学期','2020-2021 第二学期','2021-2022 第一学期'))
Item_ListBox=tk.Listbox(frm_11,listvariable=var_item)
Item_ListBox.pack()

def Get_item_cause(d_dict):
    global part_var_item
    part_var_item =Item_ListBox.get(Item_ListBox.curselection())

    if part_var_item == "2021-2022 第一学期":
        term_3 = [p3[part_var_item] for p3 in d_dict if p3.get('2021-2022 第一学期')]
        return term_3
    elif part_var_item == "2020-2021 第二学期":
        term_2 = [p2[part_var_item] for p2 in d_dict if p2.get('2020-2021 第二学期')]
        return term_2
    elif part_var_item == "2020-2021 第一学期":
        term_1 = [p1[part_var_item] for p1 in d_dict if p1.get('2020-2021 第一学期')]
        return term_1
    else:
        pass

# def show_content():
#     for ki in range(3):
#         kr_container = []
#         term = point_term(container_all_dict)  # 各学期信息
#         c_g, c_c = grade_point(term)  # 通过学期信息获取待计算的数据（用于计算平均绩点）
#         kr = con_point(c_g, c_c)  #
#         kr_container.append(kr)
#         # plt.bar(range(len(c_g)), c_g)
#         try:
#             plt.subplot(2, 2, ki + 1)
#         except Exception:
#             pass
#         else:
#             plt.grid(True)
#             plt.axis([0, 10, 0, 100])
#             plt.title(study_term_set[ki] + '的学习情况:' + str(kr)[0:4], fontsize=10)
#             plt.ylabel('成绩', color='r', fontsize=10)
#             plt.plot(range(len(c_g)), c_g, color_lab[ki], linewidth='5.0', alpha=0.5)
#
#     plt.subplot(2, 2, 4)
#     plt.grid(True)
#     plt.axis([0, 30, 0, 100])
#     plt.plot(grade, color_lab[3], linewidth='5.0', alpha=0.5)
#
#     plt.show()
def image_grade_button():
    global var_l_2
    item_PartYear=[]
    button_lb_term=Get_item_cause(container_all_dict)
    button_lb_c_g,button_lb_c_c=grade_point(button_lb_term)
    button_lb_kr=con_point(button_lb_c_g,button_lb_c_c)#绩点
    var_l_2.set('学期绩点: '+str(button_lb_kr))
    # print(button_lb_term)
    for k in button_lb_term:
        for i_key,i_values in k.items():
            item_PartYear.append(i_key)


    label=item_PartYear
    rects=plt.bar(range(len(label)),button_lb_c_g,tick_label=label)
    for rect in rects:
        height=rect.get_height()
        plt.text(rect.get_x()+rect.get_width()/2.0-0.1,height*1.01,height)
    plt.grid()
    plt.title(part_var_item)
    plt.xticks(fontsize=6)
    plt.xlabel('学科')
    plt.ylabel('成绩')
    plt.show()



button_Choose_item=tk.Button(frm_11,text='获取可视化',width=10,height=1,bg='red',command=image_grade_button)
button_Choose_item.pack()




root.mainloop()

# TODO 可视化，web ,头痛碎觉(优化），添加标题，标签，添加其他项目
