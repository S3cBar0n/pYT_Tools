# import os
import subprocess  # Using for Device Summary feature
import platform  # Using for Device Summary feature
# import pandas
# import win32api  # This has to be installed but not referenced in the program
import pyad.adquery
import pyad.aduser
import pyad.pyadutils
import utils

domain_name = "DOMAINNAME"


def main():
    print("1. AD User Summary")
    print("2. AD Device Summary")
    print("3. AD New Hire Checkup")
    print("4. Ping")
    print("5. Port Checker")
    print("0. Exit")
    try:
        selection = int(input("Select one of the numbered items: "))
    except:
        print("Please enter an integer...")
        # print(e)
        main()

    if selection <= 5:
        if selection == 1:
            print("Welcome to the AD User Summary Module")
            try:
                ad_sum()
            except Exception as e:
                print(e)
                pass

            main()
        elif selection == 2:
            print("Welcome to the AD Computer Summary Module")
            try:
                dev_sum()
            except Exception as e:
                print(e)
                pass

            main()
        elif selection == 3:
            print("Welcome to the AD New Hire Checkup Module")
            print("Why did the programmer abandon his project?")
            print("I don't know actually, you tell me")
            # ad_check()

            main()
        elif selection == 4:
            print("Welcome to the Ping Module")
            print("pong")
            # ping

            main()
        elif selection == 5:
            print("Welcome to the Port Checker Module")
            print("Shouldn't you take me out to dinner first?")
            # port_check()

            main()
        elif selection == 0:
            print("Exiting")
        else:
            print("Your selection is not listed.")
    else:
        print("Your selection is not listed.")
        main()


# Active Directory User Summary Module - returns information on a specified user
def ad_sum():
    try:
        username = input("Enter in a username: ")
        if "$" in username:
            print('To search for a device please use the AD Device Summary Module...')
            main()
        else:
            pass
    except Exception as e:
        print(e)
        print("You did not enter in a username...")
        main()

    # Launches the query service
    q = pyad.adquery.ADQuery()

    # Selects our parameters for the search and aims the search to a specific user
    q.execute_query(
        attributes=["SamAccountName", "Useraccountcontrol", "distinguishedName"],
        where_clause="SamAccountName = '{}'".format(username),
        base_dn=f"DC={domain_name}, DC=local"
    )

    # This is checking to see if the search returned any content, if it doesn't the user doesn't exist
    if q.get_row_count() <= 0:
        print(f"User {username} does not exist on {domain_name}.local")
        fail_opt = input("YES to search again, type anything else to return to menu: ")
        if fail_opt == "YES":
            ad_sum()
        else:
            main()
    else:
        pass

    # This is creating the print out for our summary and checks to see if the user is locked out
    for row in q.get_results():
        if row["Useraccountcontrol"] is not None and row["SamAccountName"] is not None:
            # if (row["Useraccountcontrol"] / 2 % 2 != 0):
            #     disabled = True
            # else:
            #     disabled = False

            # Checks to see if the user has any lockout time
            try:
                if (row['lockoutTime'] is None) or (pyad.pyadutils.convert_bigint(row['lockoutTime']) == 0):
                    lockedout = False
                else:
                    lockedout = True
            except Exception as e:
                print(e)
            try:
                description = str(row["description"])
                for sym in (("'", ""), (",", ""), ("(", ""), (")", "")):
                    description = description.replace(*sym)
                print("*************************************************")
                print(f'{row["displayName"]} ({row["SamAccountName"]})')
                print(description.replace(",", ""))
                print(row["mail"])
                print(f'Phone: {row["telephoneNumber"]}')
                print(f'Fax: {row["facsimileTelephoneNumber"]}')

                manager = (row["manager"]).split(",", 1)[0]
                print(f'Manager: {manager.replace("CN=", "")}')

                print(f"Password Set: {pyad.pyadutils.convert_datetime(row['pwdLastSet'])}")
                print(f'Locked out status: {lockedout}')
                print("*************************************************")
            except Exception as e:
                print(e)
                continue
        else:
            print(f"User {username} does not exist on {domain_name}.local")
            fail_opt = input("YES to search again, type anything else to return to menu: ")
            if fail_opt == "YES":
                ad_sum()
            else:
                break

    add_opt = input("Want to see more info? MORE for more, YES to search again, type anything else to return to menu: ")
    if add_opt == "MORE":

        print("*************************************************")
        print(f'Groups that {username} is a part of:')
        groups = sorted(row["memberOf"], key=str.lower)
        for group in groups:
            group = group.split(",", 1)[0]
            print(group.replace('CN=', ''))
        print("*************************************************")

        re_run = input("To search for another user enter YES, otherwise enter anything: ")
        if re_run == "YES":
            ad_sum()
        else:
            pass
    elif add_opt == "YES":
        ad_sum()
    else:
        pass


# Active Directory Computer Summary - Returns information on a specicified device.
def dev_sum():
    device = input("Enter Device ID: ")
    device = device + "$"

    q = pyad.adquery.ADQuery()

    # Selects our parameters for the search and aims the search to a specific user
    q.execute_query(
        attributes=["SamAccountName", "Useraccountcontrol", "cn", "distinguishedName"],
        where_clause="SamAccountName = '{}'".format(device),
        base_dn=f"DC={domain_name}, DC=local"
    )

    # This is checking to see if the search returned any content, if it doesn't the user doesn't exist
    if q.get_row_count() <= 0:
        print(f"Device {device} does not exist on {domain_name}.local")
        fail_opt = input("YES to search again, type anything else to return to menu: ")
        if fail_opt == "YES":
            dev_sum()
        else:
            main()
    else:
        pass

    # This is creating the print out for our summary and checks to see if the user is locked out
    for row in q.get_results():
        if row["Useraccountcontrol"] is not None and row["SamAccountName"] is not None:
            # if (row["Useraccountcontrol"] / 2 % 2 != 0):
            #     disabled = True
            # else:
            #     disabled = False

            try:
                print("*************************************************")
                print(f'{row["cn"]} \n')

                dev_location = str(row['distinguishedName']).replace(",OU=", ",DC=")
                dev_location = dev_location.split(',DC=')
                print("*********** Active Directory Location ***********")
                for item in dev_location[::-1]:
                    print(item)
                # print(f'Device Location: {row["distinguishedName"]}')

                utils.lookup(row["cn"])

                print("*************************************************")
            except Exception as e:
                print(e)
                continue
        else:
            print(f"User {device} does not exist on {domain_name}.local")
            fail_opt = input("YES to search again, type anything else to return to menu: ")
            if fail_opt == "YES":
                dev_sum()
            else:
                break

    add_opt = input("Want to see more info? MORE for more, YES to search again, type anything else to return to menu: ")
    if add_opt == "MORE":

        print("*************************************************")
        print(f'Groups that {device} is a part of:')
        groups = sorted(row["memberOf"], key=str.lower)
        for group in groups:
            group = group.split(",", 1)[0]
            print(group.replace('CN=', ''))
        print("*************************************************")

        re_run = input("To search for another user enter YES, otherwise enter anything: ")
        if re_run == "YES":
            dev_sum()
        else:
            pass
    elif add_opt == "YES":
        dev_sum()
    else:
        pass


main()
