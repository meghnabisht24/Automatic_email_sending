import pandas as pd
import datetime
import smtplib

def sendEmail(to, sub, msg):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('chaudharyrohan839@gmail.com', 'rohanraghav13')
    server.sendmail( 'chaudharyrohan839@gmail.com' , to ,f"Subject:{sub}\n\n {msg}")
    server.close()
    



if __name__ =="__main__":
    df = pd.read_excel("data.xlsx")

    today = datetime.datetime.now().strftime("%d-%m")
    yearnow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    
    for index, item in df.iterrows():
        bday = (item['BIRTHDAY']).strftime("%d-%m")

        if(today == bday and  yearnow not in str(item["YEAR"])):
            sendEmail( item['Email'],"HAPPY BIRTHDAY", item['DIALOGUE']) 
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i, "YEAR"]

        df.loc[i, "YEAR"] = str(yr) + ',' + str(yearnow)

    df.to_excel('data.xlsx',index= False)

         



