import os
import csv

cur = os.getcwd()

def write_opt(prefix,prod_num,vol_num,start_i,end_i):
    prod_fol = f"{prefix}_PROD{format(prod_num, '04d')}"
    vol_fol = f"VOL{format(vol_num, '05d')}"
    file1 = open(f"{cur}\\{prod_fol}\\{prefix}_PROD{format(prod_num, '03d')}.dat", "r")
    lines = file1.readlines()
    lines = lines[1:]
    opt_arr = []
    lines_arr = []
    l_index = 0
    for line in lines:
        line_arr = line.split("þ")
        # print(line_arr[1],line_arr[3])
        start = int(line_arr[1].split(f"{prefix}")[1])
        end = int(line_arr[3].split(f"{prefix}")[1])
        lines_arr.append([l_index,start,end])
        l_index +=1

    for i in range(start_i,end_i):
        i_name = format(i, '08d')
        one = f"{prefix}{i_name}"
        two = f"VOL{format(vol_num, '05d')}"
        three = f".\\{vol_fol}\\IMAGES\\{one}.jpg"
        find_head = False
        for line_arr in lines_arr:
            if i == line_arr[1]:
                number = lines[line_arr[0]].split("þ")[9]
                # print(lines[line_arr[0]].split("þ")[9])
                find_head = True
        if find_head:
            four = "Y"
            fivesix = ""
            seven = number
        else:
            four = ""
            fivesix = ""
            seven = "" 
        opt_arr.append([one,two,three,four,fivesix,fivesix,seven])

    with open(f"{cur}\\{prod_fol}\\{prefix}_PROD{format(prod_num, '03d')}.opt", "w+", newline="") as opt:
        optwriter = csv.writer(opt, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        optwriter.writerows(opt_arr)


def isInt(s):
    try: 
        int(s)
        return True
    except ValueError:
        return False

