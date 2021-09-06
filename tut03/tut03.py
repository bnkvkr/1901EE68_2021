import os
path2 = './output_by_subject'
isfile2 = os.path.isdir(path2)
if(isfile2 == False):
    os.mkdir(path2)

path1 = './output_individual_roll'
isFile = os.path.isdir(path1)
if(isFile == False):
    os.mkdir(path1)


def output_by_subject():
    with open('regtable_old.csv', 'r') as ff:
        for line in ff:
            words = line.split(',')
            subno = words[3]
            with open("./output_by_subject/"+subno+".csv", 'a+') as ap:

                if ap.tell() == 0:
                    ap.writelines(["rollno,register_sem,subno,sub_type\n"])
                else:

                    with open("./output_by_subject/"+subno+".csv", 'r') as f:
                        all_lines = f.readlines()
                        new_data = words[0] + "," + words[1] + \
                            "," + words[3]+"," + words[8]
                        if new_data not in all_lines:
                            ap.writelines([new_data])

    return


def output_individual_roll():

    with open('regtable_old.csv', 'r') as ff:
        for line in ff:
            words = line.split(',')
            roll = words[0]
            with open("./output_individual_roll/"+roll+".csv", 'a+') as ap:

                if ap.tell() == 0:
                    ap.writelines(["rollno,register_sem,subno,sub_type\n"])
                else:

                    with open("./output_individual_roll/"+roll+".csv", 'r') as f:
                        all_lines = f.readlines()
                        new_data = words[0] + "," + words[1] + \
                            "," + words[3]+"," + words[8]
                        if new_data not in all_lines:
                            ap.writelines([new_data])
    return


output_individual_roll()
output_by_subject()
