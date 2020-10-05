from functions import *


log_into_dxc()

counter = 1

while True:

    try:

        temp = find_line_attributes(counter)[1]

        if "DES Entry Request" in temp:

            find_des_number(counter)

            counter = counter - 1

        else:

            pass

        counter += 1

    except NoSuchElementException:

        refresh_page()
        counter = 1
