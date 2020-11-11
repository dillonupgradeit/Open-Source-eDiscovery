import os
cur = os.getcwd()
def write_dat(prefix,vol_num,prod_num):
    prod_fol = f"{prefix}_PROD{format(prod_num, '04d')}"
    file1 = open(f"{cur}\\temp\\{prefix}_PRO{format(prod_num, '03d')}.dat", "r")
    lines = file1.readlines()
    line1 = lines[0]
    lines = lines[1:]
    lines.sort()
    writer = open(f"{cur}\\{prod_fol}\\{prefix}_PROD{format(prod_num, '03d')}.dat", "w")
    writer.write(f"{line1}")
    for line in lines:
        writer.write(f"{line}")
    writer.close()

