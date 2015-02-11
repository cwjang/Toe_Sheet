import matplotlib.pyplot as plt
plt.rcdefaults()
Plot_list = []

def Parse_info(f):
    for l in f:
        if 'Start' in l:
            break
    for l in f:
        if l and len(l) > 5:
            title, ylabel, x_names, y_vals, error, x_pos, color_list = l.rstrip('\n').split('>')[:7]
            list0 = []
            for data in [x_names, y_vals, error, x_pos, color_list]:
                list0.append(data.split(',')[:-1])
            name_list = []
            T = '************'
            for name in list0[0]:
                n = name.replace('\\n', '\n')
                if T == n.split('\n')[0]:
                    n = n.lstrip(T)
                else:
                    if n != '':
                        T = n.split('\n')[0]
                name_list.append(n)
            list1 = []
            for data in list0[1: -1]:
                list1.append([float(i) for i in data])
            Plot_list.append([title, ylabel] + [name_list] + list1 + [list0[-1]])
                

def Plotting(title, ylabel, x_names, y_vals, error, x_pos, color_list):
    fig = plt.figure()
    plt.bar(x_pos, y_vals, yerr = error, align = 'center', color = color_list, alpha = 0.6, ecolor = 'black')
    plt.xticks(x_pos, x_names)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.draw()
    fig.show()

File = 'C:/Users/cwj/Desktop/Plotting_info.txt'

#File = 'E:/Data/qPCR/Plotting_info_p53nullMEF_exp.txt'

with open(File, 'r') as f:
    Parse_info(f)
for title, ylabel, x_names, y_vals, error, x_pos, color_list in Plot_list:
    Plotting(title, ylabel, x_names, y_vals, error, x_pos, color_list)
