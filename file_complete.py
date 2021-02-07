from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import tkinter as tk
from PIL import Image, ImageTk


# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1pR8Odx76FEjuIGCkIbd9WszrTfyfbVdrE5HovsgsaSE'
SAMPLE_RANGE_NAME = 'Sheet1'

master = tk.Tk()
var = tk.IntVar()

def get_data():
    global E
    E = E_input.get()
    print(E)
    var.set(1)

label = tk.Label(master, text="Code Number")
label.grid(row=0)
E_input = tk.Entry(master)
E_input.grid(row=0, column=1)
button = tk.Button(master, text='Submit', command=get_data)
button.grid(row=3, column=1, pady=4)

print("waiting...")
button.wait_variable(var)
print("done waiting.")


def main():
    global date
    date = []
    global office
    office = []
    global person
    person = []
    global comment
    comment = []
    row_number = -1
    print(E)

    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    for row in values:
        row_number = row_number + 1
        if row[1] == E:
            code = row[1]
            try:
                for i in range(3,20):
                    date.append(row[i])
            except:
                pass
            row = values[row_number+1]
            subject = row[1]
            try:
                for i in range(3,20):
                    office.append(row[i])
            except:
                pass
            row = values[row_number+2]
            requester = row[1]
            try:
                for i in range(3,20):
                    person.append(row[i])
            except:
                pass
            row = values[row_number+3]
            project_title = row[1]
            try:
                for i in range(3,20):
                    comment.append(row[i])
            except:
                pass
            row = values[row_number+4]
            mobile = row[1]
            break
    global basic_text
    basic_text = ("Code: " + str(code) + "\nSubject: " + str(subject) + "\nRequester: " + str(requester) + "\nProject Title: " + str(project_title) + "\nMobile: " + str(mobile))
    for item in date:
        print ("Date: " + item)
    for item in person:
        print ("Person: " + item)
    for item in office:
        print ("Office: " + item)
    for item in comment:
        print ("Comment: " + item)
        

if __name__ == '__main__':
    try:
        main()
        label.grid_forget()
        E_input.grid_forget()
        button.grid_forget()

        load = Image.open("download.png")
        load = load.resize((150,113))
        render = ImageTk.PhotoImage(load)
        img_left = tk.Label(master, image=render)
        img_left.image = render
        img_left.grid(column = 0, row = 0)

        load2 = Image.open("GrantsOffice.png")
        load2 = load2.resize((150,113))
        render2 = ImageTk.PhotoImage(load2)
        img_right = tk.Label(master, image=render2)
        img_right.image = render2
        img_right.grid(column = 2, row = 0)

        basic_data1 = tk.Label(master, text = basic_text)
        basic_data1.config(font=("Courier", 24))
        basic_data1.grid(column = 1, row=2)
            

        def show_changes():
            basic_data1.grid_forget()
            divider = tk.Label(master, text = "\nChanges:", font=("Courier", 24, 'bold'))
            divider.grid(column = 1, row=2)

            table_top = tk.Label(master, text = "Date   Office   Person   Comment", font=("Courier", 24, 'bold'))
            table_top.grid(column = 1, row=3)

            for i in range(0, len(date)):
                table_text = (date[i] + "   " + office[i] + "   " + person[i] + "   " + comment[i])
                tk.Label(master, text = table_text, font=("Courier", 24)).grid(column = 1, row = (i+4))
            next_button.grid_forget()


        next_button = tk.Button(master, text='Next', command=show_changes, width = 15)
        next_button.grid(column = 1, row=1)

    except:
        label.config(text = "Data Not Found")
        nth = input()
        



