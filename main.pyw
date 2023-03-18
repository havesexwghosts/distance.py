import pandas as pd, time, datetime, threading
from customtkinter import *
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class application():
    def __init__(self):
        threading.Thread(target=self.main).start()
    
    def openfile(self):
        filetypes = (('Excel Files', '*.xlsx'), ('All Files', '*.*'))
        self.filename = filedialog.askopenfilename(title='Select excel file', initialdir='/', filetypes=filetypes)
        self.selected_file.configure(text="File selected {}".format(self.filename))
        dataframe = pd.read_excel(self.filename)
        self.total_distances.configure(text=f"Total distances {len(dataframe)}")
    
    def openfolder(self):
        self.folderdirectory = filedialog.askdirectory(title='Select output location', initialdir='/')
        self.selected_folder.configure(text="Folder selected {}".format(self.folderdirectory))

    def run(self):
        threading.Thread(target=self.calculate).start()
        threading.Thread(target=self.timer).start()

    def calculate(self):
        self.running = True
        self.calculated = 0
        self.errors = 0
        self.driver = webdriver.Chrome()
        self.driver.get('https://www.google.com/maps/dir///@0,0,15z/data=!4m2!4m1!3e0')
 
        self.dataframe = pd.read_excel(self.filename)
        for row in range(len(self.dataframe)):
            if pd.isna(self.dataframe.loc[row, 'Result']) or self.dataframe.loc[row, 'Result'] == 'Error':
                if pd.isnull(self.dataframe.loc[row, 'Origem']) is not True and pd.isnull(self.dataframe.loc[row, 'Destino']) is not True and self.dataframe.loc[row, 'Origem'] != self.dataframe.loc[row, 'Destino']:
                    origem  = self.dataframe.loc[row, 'Origem'] 
                    destino = self.dataframe.loc[row, 'Destino']
                    try:
                        self.driver.find_element('xpath', '//*[@id="sb_ifc50"]/input').clear()
                        self.driver.find_element('xpath', '//*[@id="sb_ifc51"]/input').clear()

                        self.driver.find_element('xpath', '//*[@id="sb_ifc50"]/input').send_keys(origem)
                        self.driver.find_element('xpath', '//*[@id="sb_ifc51"]/input').send_keys(destino)
                        self.driver.find_element('xpath', '//*[@id="directions-searchbox-1"]/button[1]').click()
                        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(('xpath', '//*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div')))
                        self.distance = str(self.driver.find_element('xpath', '//*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div').text[:-3])
                        self.dataframe.loc[row, 'Result'] = int(self.distance.replace('.', ''))
                        self.calculated += 1
                        self.total_searches.configure(text=f"Searches Made: {self.calculated}")
                    except Exception as e:
                        print(e)
                        self.dataframe.loc[row, 'Result'] = ''
                        self.errors += 1
                        self.total_errors.configure(text=f"Unable to fetch distance: {self.errors}")
                else:
                    print('Field is empty or origem and destino match the same value.')
                    self.dataframe.loc[row, 'Result'] = 0
            else:
                print("No values to calculate")
        self.dataframe.to_excel(self.folderdirectory + '/' + time.strftime('%H.%M.%S') + '.xlsx', index=False)
        self.driver.close()
        self.running = False
        return
    
    def timer(self):
        t = 0
        while self.running:
            t += 1
            secs = t % 60
            mins = (t// 60) % 60
            formated = datetime.time(minute=mins, second=secs)
            self.total_time.configure(text=f'{formated.strftime("%M:%S")}')
            time.sleep(1)
        return

    def main(self):
        self.root = CTk()
        self.root.geometry("1100x580")
        self.root.resizable(False, False)
        self.root.title("V2")

        self.sidebarframe1 = CTkFrame(self.root, height=580, width=280, corner_radius=0)
        self.sidebarframe1.place(x=820, y=0)

        self.total_distances = CTkLabel(self.sidebarframe1, text="Total searches: ", font=("Roboto",15))
        self.total_distances.place(x=10, y=10)

        self.total_searches = CTkLabel(self.sidebarframe1, text="Searches made: ", font=("Roboto",15))
        self.total_searches.place(x=10, y=60)

        self.total_errors = CTkLabel(self.sidebarframe1, text="Unable to fetch distance: ", font=("Roboto",15))
        self.total_errors.place(x=10, y=110)

        self.total_time = CTkLabel(self.sidebarframe1, text="Time elapsed: ", font=("Roboto",15))
        self.total_time.place(x=10, y=160)

        self.execute = CTkButton(self.sidebarframe1, text="Execute", height=30, width=150, command=lambda : self.run())
        self.execute.place(x=10, y=200)

        self.frame1 = CTkFrame(self.root, width=780, height=520)
        self.frame1.place(x=20, y=20)

        self.label1 = CTkLabel(self.frame1, text="Google Maps Selenium Bot", font=("Roboto", 24))
        self.label1.place(x=10, y=30)

        self.selected_file = CTkLabel(self.frame1, text="File selected: ", font=("Roboto",18))
        self.selected_file.place(x=10, y=100)
        self.open_file = CTkButton(self.frame1, text="Open file", height=30, width=150, command=lambda : self.openfile())
        self.open_file.place(x=10, y=130)

        self.selected_folder = CTkLabel(self.frame1, text="Output location: ", font=("Roboto", 18))
        self.selected_folder.place(x=10, y=180)
        self.open_folder = CTkButton(self.frame1, text="Open folder", height=30, width=150, command=lambda : self.openfolder())
        self.open_folder.place(x=10, y=210)

        self.label2 = CTkLabel(self.frame1, text="Made by Watanabe, V2",  font=("Roboto",15))
        self.label2.place(x=10, y=430)

        self.label3 = CTkLabel(self.frame1, text="Updated the GUI with customtkinter. Use of threading module and minor bug fixes.",  font=("Roboto",12))
        self.label3.place(x=10, y=460)

        self.label4 = CTkLabel(self.frame1, text="GitHub repo: https://github.com/6ixess/distance.py.",  font=("Roboto",12))
        self.label4.place(x=10, y=490)

        self.root.mainloop()

if __name__ == "__main__":
    application()