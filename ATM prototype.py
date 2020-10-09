import xlrd
import xlwt
from xlutils.copy import copy
import smtplib
server=smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login('teatimecomedy@gmail.com','saai saai')
file='/home/rajkumar/Desktop/machine_learning_training/ATM_data.xlsx'
print('Welcome To Techienest Bank ATM!')
while(1):
    print('Enter 1 to login as Admin')
    print('Enter 2 to login as USER')
    print('Enter 3 to Exit')
    s=int(input())
    if(s==1):
        workbook=xlrd.open_workbook(file)
        sheet=workbook.sheet_by_index(1)
        usernames=[sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
        while(1):
            username=input('Enter the adminname:')
            if username in usernames:
                m=usernames.index(username)
                password=input('Enter the password:')
                if(password==sheet.cell_value(m+1,1)):
                    print('Admin Login Sucessful!Manage Your Customers\n')
                    while(1):
                        print('1.Add Customer\n2.Delete Customer\n3.View All Customers details\n4.exit\n')
                        ab=int(input('Enter Your Choice:'))
                        if(ab==1):
                            print('Enter the details of customer below\n')
                            uname=input('Enter the username:')
                            upass=input('Enter the password:')
                            uamount=float(input('Enter the initial balance of the user:'))
                            umail=input('enter the mail_id of the user:')
                            workbook1=xlrd.open_workbook(file)
                            sheet1=workbook1.sheet_by_index(0)
                            rd=copy(workbook1)
                            d=rd.get_sheet(0)
                            d.write((sheet1.nrows)+1,0,uname)
                            d.write((sheet1.nrows+1),1,upass)
                            d.write((sheet1.nrows)+1,2,uamount)
                            d.write((sheet1.nrows)+1,3,umail)
                            rd.save(file)
                            print('customer details added sucessfully\n')
                        elif(ab==2):
                            workbook1=xlrd.open_workbook(file)
                            sheet1=workbook1.sheet_by_index(0)
                            usernames=[sheet1.cell_value(k,0) for k in range(1,sheet1.nrows)]
                            ur=input(('Enter the username of the to delete from data base:'))
                            if ur in usernames:
                                mr=usernames.index(ur)
                                mr=mr+1
                                rd=copy(workbook1)
                                d=rd.get_sheet(0)
                                d.write(mr,0,'')
                                d.write(mr,1,'')
                                d.write(mr,2,'')
                                d.write(mr,3,'')
                                rd.save(file)
                                print('Deletion of customer from the database is sucessful.\n')
                            else:
                                print('entered customer is not in the list')
                        elif(ab==3):
                            print('customers details details\n')
                            workbook1=xlrd.open_workbook(file)
                            sheet1=workbook1.sheet_by_index(0)
                            for count in range(sheet1.nrows):
                                print(sheet1.row_values(count))
                        elif(ab==4):
                            break
                    print('adminLoginsucessfull')
                    break
                else:
                    print('Check Your Password')
            else:
                
                print('Please Check Your Name')

    elif(s==2):
        workbook1=xlrd.open_workbook(file)
        sheet1=workbook1.sheet_by_index(0)
        usernames=[sheet1.cell_value(k,0) for k in range(1,sheet1.nrows)]
        while(1):
            username=input('Enter the user name:')
            if username in usernames:
                m=usernames.index(username)
                password=input('Enter the password:')
                if(password==sheet1.cell_value(m+1,1)):
                    print('Loginsucessfull Enjoy Our Facilities '+username)
                    print('1.Check Your Balance\n2.Withdraw\n3.deposit\n4.Change Password\n')
                    b=int(input('enter your choice:'))
                    if(b==1):
                        bal=str(sheet1.cell_value(m+1,2))
                        print('Your Account Balance:'+bal)
                        server.sendmail('teatimecomedy@gmail.com',sheet1.cell_value(m+1,3),'This is a Mail From Techienest Bank !\nGenerated due to your action at ATM to Check Yourbalance.\n Your Balance is:RS '+str(bal))
                    elif(b==2):
                        rd=copy(workbook1)
                        d=rd.get_sheet(0)
                        bal=sheet1.cell_value(m+1,2)
                        wd=float(input('Enter the amount you want to withdraw:'))
                        change=bal-wd
                        d.write(m+1,2,change)
                        rd.save(file)
                        print('Amount Withdrawn Sucessful!\n')
                        server.sendmail('teatimecomedy@gmail.com',sheet1.cell_value(m+1,3),'This is a mail from Techienest bank!\n to intimate that Your account has been debited with amount '+str(wd)+'\nupdated balance is '+str(change))
                    elif(b==3):
                        rd=copy(workbook1)
                        d=rd.get_sheet(0)
                        bal=sheet1.cell_value(m+1,2)
                        cd=float(input('Enter the amount you want to deposit:'))
                        change=bal+cd
                        d.write(m+1,2,change)
                        rd.save(file)
                        print('Amount deposited Sucessfuly!\n')
                        server.sendmail('teatimecomedy@gmail.com',sheet1.cell_value(m+1,3),'This is a mail from Techienest bank!\n To intimate that Your account has been deposited with amount '+str(cd)+'\nupdated balance is '+str(change))
                    elif(b==4):
                        while(1):
                            pass1=input(('Enter Your old Password:'))
                            if(pass1==sheet1.cell_value(m+1,1)):
                            
                                while(1):
                                    pass2=input('Enter Your New Password:')
                                    pass3=input('Enter Your New Password again:')
                                    if(pass2==pass3):
                                        rd=copy(workbook1)
                                        d=rd.get_sheet(0)
                                        d.write(m+1,1,pass2)
                                        rd.save(file)
                                        print('Password Changed Successful!Now you Can Login With New Password\n')
                                        server.sendmail('teatimecomedy@gmail.com',sheet1.cell_value(m+1,3),'This is the mail from Techienest Bank to intimate you that your password has been changed')
                                        break
                                    else:
                                        print('Password do not match!')
                                break
                            else:
                                print('Your Old Password is Wrong Try again!')
                    else:
                        print('Wrong Entry Login again\n')
                    break
                else:
                    print('Check Your Password')
            else:
                
                print('Please Check Your username')
    elif(s==3):
        break
    else:
        print('Wrong Entry try Again')
