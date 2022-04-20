from tkinter import *
from PIL import ImageTk, Image
from pio_excel_formater import *
from tkinter import ttk
from datetime import date
import operator
import re

#1abc9c

global login_var, users, today, switch_func, switch_func2, switch_func3, user_selected

login_var = "off"

users =[["rastko", "rg5x5"]]

today = date.today()
today =  today.strftime("%d.%m.%y")


switch_func=0
switch_func2=0
switch_func3=0

user_selected = "off"

#Creating main window
root = Tk()
root.title("PIO EVIDENCIJA 2022")
root.iconbitmap("img_resources/pio.ico")
root.resizable(width=False, height=False)


#Defining Images for backgound and buttons
add_img = ImageTk.PhotoImage(Image.open("img_resources/add.png"))
diskette_img = ImageTk.PhotoImage(Image.open("img_resources/diskette.png"))
diskette_small_img = ImageTk.PhotoImage(Image.open("img_resources/diskette_small.png"))
search_img = ImageTk.PhotoImage(Image.open("img_resources/search.png"))
money_img = ImageTk.PhotoImage(Image.open("img_resources/money.png"))
stats_img = ImageTk.PhotoImage(Image.open("img_resources/stats.png"))
ok_img = ImageTk.PhotoImage(Image.open("img_resources/ok.png"))
cancel_img = ImageTk.PhotoImage(Image.open("img_resources/cancel.png"))
store_img = ImageTk.PhotoImage(Image.open("img_resources/store.png"))
more_img = ImageTk.PhotoImage(Image.open("img_resources/more.png"))
search_img_small = ImageTk.PhotoImage(Image.open("img_resources/search24x24.png"))
name_img_small = ImageTk.PhotoImage(Image.open("img_resources/name_small.png"))
money_img_small = ImageTk.PhotoImage(Image.open("img_resources/money_small.png"))
receipt_img_small = ImageTk.PhotoImage(Image.open("img_resources/receipt_small.png"))
question_img_small = ImageTk.PhotoImage(Image.open("img_resources/question_small.png"))
trash_img_small = ImageTk.PhotoImage(Image.open("img_resources/trash_small.png"))
edit_img_small = ImageTk.PhotoImage(Image.open("img_resources/edit_small.png"))
pay_img_small = ImageTk.PhotoImage(Image.open("img_resources/pay_small.png"))
pay2_img_small = ImageTk.PhotoImage(Image.open("img_resources/pay2_small.png"))
enlarge_img_small = ImageTk.PhotoImage(Image.open("img_resources/enlarge_small.png"))
cancel_pay_img_small = ImageTk.PhotoImage(Image.open("img_resources/cancel_pay_small.png"))
cross_img_small = ImageTk.PhotoImage(Image.open("img_resources/cross_small.png"))
warning_img = ImageTk.PhotoImage(Image.open("img_resources/warning.png"))
print_img_small = ImageTk.PhotoImage(Image.open("img_resources/print_small.png"))
money_img_small = ImageTk.PhotoImage(Image.open("img_resources/money_small.png"))
pillars_img = ImageTk.PhotoImage(Image.open("img_resources/pillars.png"))
add_payment_img = ImageTk.PhotoImage(Image.open("img_resources/add_payment.png"))
question_img = ImageTk.PhotoImage(Image.open("img_resources/question.png"))
help_img = ImageTk.PhotoImage(Image.open("img_resources/help.jpg"))
calc_img = ImageTk.PhotoImage(Image.open("img_resources/calc.png"))
calc_small_img = ImageTk.PhotoImage(Image.open("img_resources/calc_small.png"))
late_user_img = ImageTk.PhotoImage(Image.open("img_resources/late_user.png"))
unknown_payments_img = ImageTk.PhotoImage(Image.open("img_resources/unknown_payments.png"))
payments_date_img = ImageTk.PhotoImage(Image.open("img_resources/payments_date.png"))
skipped_payments_img = ImageTk.PhotoImage(Image.open("img_resources/skipped_payments.png"))
excel_img = ImageTk.PhotoImage(Image.open("img_resources/excel.png"))
arrow_right_img = ImageTk.PhotoImage(Image.open("img_resources/arrow_right.png"))
arrow_left_img = ImageTk.PhotoImage(Image.open("img_resources/arrow_left.png"))
wide_view_img = ImageTk.PhotoImage(Image.open("img_resources/wide_view.png"))
narrow_view_img = ImageTk.PhotoImage(Image.open("img_resources/narrow_view.png"))
return_img = ImageTk.PhotoImage(Image.open("img_resources/return.png"))
note_img = ImageTk.PhotoImage(Image.open("img_resources/note.png"))



global store_index
store_index = StringVar()
store_index.set("0")


def login_page():

	global login_image, login_background, user, password, login_frame

	login_frame = Frame(root, width=613, height=560)
	login_frame.grid(row=0, column=0)

	login_background = Canvas(login_frame, width=640, height=674)
	login_background.pack(fill=BOTH, expand=1)

	login_img = Image.open("img_resources/background.jpg")
	login_img = login_img.resize((640, 674), Image.ANTIALIAS)

	login_image = ImageTk.PhotoImage(login_img)
	login_background.create_image(0, 0, image=login_image, anchor=NW)
	login_background.create_image(193, 120, image=pillars_img, anchor=NW)
	login_background.create_text(318, 70, text="PIO EVIDENCIJA 2022", font=("Courier", 36, "bold"))
	login_background.create_text(248, 400, text="Korisničko ime:")
	login_background.create_text(229, 450, text="Lozinka:")

	user = Entry(login_background, width=37)
	user_position = login_background.create_window(320, 415, window=user)

	password = Entry(login_background, width=37, show="*")
	password_position = login_background.create_window(320, 465, window=password)

	login_button = Button(login_background, text="Uloguj se", width=31, command=login)
	login_button_position = login_background.create_window(320, 515, window=login_button)

def login():

	global login_var

	login_var = "on"

	def delete_login_warning():
		login_background.delete("upozorenje")

	if user.get() not in users[0] or password.get() not in users[0]:
		login_background.create_text(300, 550, text="Pogrešna šifra ili lozinka", fill="red", tag="upozorenje")
		login_background.after(2000, delete_login_warning)
		user.delete(0, END)
		password.delete(0, END)
	
	else:
		login_background.delete("upozorenje")
		login_background.create_text(300, 550, text="Baaavoooo", fill="red", tag="upozorenje")
		login_frame.destroy()


		global background_img, background_frame, toolbar, background_image, new_user, search, save_changes, stats, payments, background

		
		background_img = Image.open("img_resources/background.jpg")
		background_img = background_img.resize((640, 560), Image.ANTIALIAS)

		#Defining top frame where main buttons will go

		toolbar = Frame(root, width=613, height=114)
		toolbar.grid(row=0, column=0)


		#Defining bottom frame with canvas background
		background_frame = Frame(root, width=615, height=560)
		background_frame.grid(row=1, column=0)

		background = Canvas(background_frame, width=640, height=560)
		background.pack(fill=BOTH, expand=1)

		background_image = ImageTk.PhotoImage(background_img)
		background.create_image(0, 0, image=background_image, anchor=NW)

		new_user = Button(toolbar, text= "Novi korisnik", fg="black", width=120, height=110, image= add_img, compound=TOP, command=add_user)
		new_user.grid(row=0, column=0)

		search = Button(toolbar, text= "Pretraga", fg="black", width=120, height=110, image= search_img, compound=TOP, command=search_user)
		search.grid(row=0, column=1)

		payments = Button(toolbar, text= "Rad sa uplatama", fg="black", width=120, height=110, image= money_img, compound=TOP, command=payments_menu)
		payments.grid(row=0, column=2)

		stats = Button(toolbar, text= "Statistika", fg="black", width=120, height=110, image= stats_img, compound=TOP, command=stats_menu)
		stats.grid(row=0, column=3)
		
		save_changes = Button(toolbar, text= "Sačuvaj izmene", fg="black", width=120, height=110, image= diskette_img, compound=TOP, command=save_document_excel)
		save_changes.grid(row=0, column=4)

def reset_param():

	global background, switch_func, switch_func2, switch_func3, user_selected

	#Reseting canvas background

	background.delete('all')
	background.create_image(0, 0, image=background_image, anchor=NW)

	#Reseting swithc variable used inside of the search user function. Used to show/hide makeshift dropdown menu
	switch_func = 0
	switch_func2 = 0
	switch_func3 = 0
	user_selected = "off"




# Function to add new user. Command is initiated on a button, content is saved inside of an excel file
def add_user():

	reset_param()


	#Defining globals for saving user data in an excel table

	global name, address, city, JMBG, credit, instalments, instalment_info, receipt, store
	
	'''
	Defining a function that takes entries from credit and instalment entries, automaticaly calculates the value of instalment,
	and enters that data into instalment_info entry box
	'''

	
	def click_on_box(event):

		credit_data = credvar.get()

		if credit_data == "":
			pass

		else:
			credit_data = credit_data.replace(",","")
			credit_data = float(credit_data)
			credit.delete(0, END)
			credit.insert(0, credit_data)

	
	def click_from_box(event):

		credit_data = credvar.get()
		
		pattern = r"^[1-9]\d*(\.\d+)?$"


			
		if credit_data == "":
			pass

		else:
			match = re.match(pattern, credit_data)
			if match:
				background.delete("upozorenje")
				credit_data = credit_data.replace(",","")
				credit_data = float(credit_data)
				credit_data = str(f'{credit_data:,}')
				credit.delete(0, END)
				credit.insert(0, credit_data)
			else:
				credit.delete(0, END)
				background.delete("upozorenje")
				background.create_text(325, 482, text="Iznos kredita mora biti upisan u odgovarajućem formatu.", fill="red", tag="upozorenje")
	

	def calculate_instalment(*args):

		credit_data = credvar.get()
		instalment_data = instalvar.get()

		pattern = r"^[1-9]\d*(\.\d+)?$"
		pattern2 = r"^[0-9]*$"

		if instalment_data == "" or credit_data == "":
			instalment_info.delete(0, END)

		else:
			credit_data = credit_data.replace(",","")
			match = re.match(pattern, credit_data)
			match2 = re.match(pattern2, instalment_data)
			if match and match2:
				background.delete("upozorenje")
				credit_data = float(credit_data)
				intalment_data = int(instalment_data)
				if int(instalment_data) > 12 or int(instalment_data) == 0:
					instalments.delete(0, END)
					instalment_info.delete(0, END)
					background.delete("upozorenje")
					background.create_text(325, 482, text="Najveći dozvoljeni broj rata je 12.", fill="red", tag="upozorenje")
				else:
					result = round(credit_data/intalment_data, 2)
					resultvar.set(result)
					instalment_info.delete(0, END)
					instalment_info.insert(0, str(f'{result:,}'))
			else:
				instalments.delete(0, END)
				instalments.delete(0, END)
				background.delete("upozorenje")
				background.create_text(325, 482, text="Kredit i rate moraju biti upisani kao brojevi.", fill="red", tag="upozorenje")
	
	#Defining description labels

	background.create_text(155, 50, text="Ime i prezime:")
	background.create_text(155, 97, text="Adresa:")
	background.create_text(155, 144, text="Grad:")
	background.create_text(155, 191, text="JMBG:")
	background.create_text(155, 238, text="Iznos kredita:")
	background.create_text(155, 285, text="Broj rata:")
	background.create_text(155, 332, text="Iznos rate:")
	background.create_text(155, 379, text="Broj računa:")
	background.create_text(155, 426, text="Maloprodaja:")	

	#Defining entry boxes for each description label where user data will be entered


	credvar = StringVar()
	instalvar = StringVar()
	resultvar = StringVar()

	name = Entry(background, text="Ime i prezime:", width=40)
	name_position = background.create_window(385, 50, window=name)

	address = Entry(background, text="Adresa:", width=40)
	address_position = background.create_window(385, 97, window=address)

	city = Entry(background, text="Grad:", width=40)
	city_position = background.create_window(385, 144, window=city)

	JMBG = Entry(background, text="JMBG:", width=40)
	JMBG_position = background.create_window(385, 191, window=JMBG)

	credit = Entry(background, text="Iznos kredita:", width=40, textvariable=credvar)
	credit_position = background.create_window(385, 238, window=credit)

	instalments = Entry(background, text="Broj rata:", width=40, textvariable=instalvar)
	instalments_position = background.create_window(385, 285, window=instalments)

	instalment_info = Entry(background, text="Iznos rate:", width=40, textvariable=resultvar)
	instalment_info_position = background.create_window(385, 332, window=instalment_info)

	receipt = Entry(background, text="Broj računa:", width=40)
	receipt_position = background.create_window(385, 379, window=receipt)

	store = Entry(background, text="Maloprodaja:", width=40)
	store_position = background.create_window(385, 426, window=store)

	credit.delete(0, END)
	instalments.delete(0, END)
	instalment_info.delete(0, END)


	credvar.trace("w", calculate_instalment)
	instalvar.trace("w", calculate_instalment)
	resultvar.trace("w", calculate_instalment)

	credit.bind("<FocusOut>", click_from_box)
	credit.bind("<FocusIn>", click_on_box)


	# Defining two buttons inside of add_user function (Save and Cancel)

	confrim_button = Button(background, text="Sačuvaj" + " "*10, fg="black", width=150, height=30, image=ok_img, compound=RIGHT, command=confirm_add_user)
	confrim_button_position = background.create_window(98, 520, window=confrim_button)

	cancel_button = Button(background, text="Nazad" + " "*10, fg="black", width=150, height=30, image=cancel_img, compound=RIGHT, command=cancel_add_user)
	cancel_button_position = background.create_window(542, 520, window=cancel_button)

	name.delete(0, END)
	address.delete(0, END)
	city.delete(0, END)
	JMBG.delete(0, END)
	credit.delete(0, END)
	instalments.delete(0, END)
	instalment_info.delete(0, END)
	receipt.delete(0, END)
	store.delete(0, END)

# Function used to confirm and save data enterd inside of add_user function 

def confirm_add_user():

	def delete_warning1():
		background.delete("warning1")

	'''
	Bunch of IF statements to check if any enty field has been left empty.
	If a field is left empty, the user will be shown a warning sign with red letters
	'''
	if instalment_info.get() == "":
		pass
	
	else:

		float_instalment_info = instalment_info.get().replace(",","")
		float_instalment_info = float(float_instalment_info)

		float_credit = credit.get().replace(",","")
		float_credit = float(float_credit)

		int_instalments = instalments.get().replace(",","")
		int_instalments = int(int_instalments)



	if name.get() == "" or address.get() == "" or city.get() == "" or JMBG.get() == "" or credit.get() == "" or instalments.get() == "" or instalment_info.get() == "" or receipt.get() == "" or store.get() == "" or float_credit < 1 or int_instalments < 1 or float_instalment_info < 1:
		background.delete("warning1")
		background.delete("upozorenje")
		background.create_text(325, 482, text="Sva polja moraju biti popunjena.", fill="red", tag="upozorenje")

	#If all fields have a value greater, or value equal or greater than 1 that means that provided user data can be saved inside of an excel file		

	else:

		update_excel_sheet(name, JMBG, address, city, credit, instalments, instalment_info, receipt, store)

		'''
		All entry fields are reset to be empty, allowing for new data to be entered
		User is notified that new customer data has been accepted with a message with red letters
		'''
		background.delete("upozorenje")
		background.create_text(325, 482, text="Korisnik upisan u bazu korisnika.", fill="red", tag="warning1")		
		background.after(5000, delete_warning1)

		name.delete(0, END)
		address.delete(0, END)
		city.delete(0, END)
		JMBG.delete(0, END)
		credit.delete(0, END)
		instalments.delete(0, END)
		instalment_info.delete(0, END)
		receipt.delete(0, END)
		store.delete(0, END)
	



#Function used to cancel data enterd inside of add_user function, all fields are cleared, and canvas background is reset to default, allowing for new data to be enterd
def cancel_add_user():
	
	reset_param()	

	name.delete(0, END)
	address.delete(0, END)
	city.delete(0, END)
	JMBG.delete(0, END)
	credit.delete(0, END)
	instalments.delete(0, END)
	instalment_info.delete(0, END)
	receipt.delete(0, END)
	store.delete(0, END)

# Function that initiates a makeshift dropdown menu. The menu offers 4 different buttons for 4 different methods of browsing trough customer data
def search_user():

	store_index.set("0")
	#Defining a switcher used to define if dropdown menu will open or close.

	global switch_func, switch_func2, switch_func3

	switch_func2 = 0
	switch_func3 = 0

	# If switch variable is 0, the menu will open
	if switch_func == 0:

		background.delete('all')
		background.create_image(0, 0, image=background_image, anchor=NW)

		#Defining 2 different search buttons 

		button1 = Button(background, text="Po radnji", fg="black", width=120, height=110, image= store_img, compound=TOP, command=show_users_by_store)
		button1_position = background.create_window(194,57, window=button1)

		button2 = Button(background, text="Drugi kriterijum", fg="black", width=120, height=110, image= more_img, compound=TOP, command=other_search_criteria)
		button2_position = background.create_window(194,174, window=button2)

		switch_func = 1

	#If the menu is 1 the menu will close	
	else:

		reset_param()

def stats_menu():

	store_index.set("0")
	#Defining a switcher used to define if dropdown menu will open or close.

	global switch_func, switch_func2, switch_func3

	switch_func = 0
	switch_func3 = 0

	# If switch variable is 0, the menu will open
	if switch_func2 == 0:

		background.delete('all')
		background.create_image(0, 0, image=background_image, anchor=NW)

		#Defining 2 different search buttons 

		stats_button1 = Button(background, text="Struktura duga", fg="black", width=120, height=110, image=calc_img, compound=TOP, command=debt_structure)
		stats_button1_position = background.create_window(450,57, window=stats_button1)

		stats_button2 = Button(background, text="Preminuli korisnici", fg="black", width=120, height=110, image=late_user_img, compound=TOP, command=sort_by_late_users)
		stats_button2_position = background.create_window(450,174, window=stats_button2)

		stats_button3 = Button(background, text="Statistika po radnjama", fg="black", width=120, height=110, image=store_img, compound=TOP, command=stats_by_store)
		stats_button3_position = background.create_window(450,291, window=stats_button3)

		switch_func2 = 1

	#If the menu is 1 the menu will close	
	else:

		reset_param()


def payments_menu():

		
	store_index.set("0")
	#Defining a switcher used to define if dropdown menu will open or close.

	global switch_func, switch_func2, switch_func3

	switch_func = 0
	switch_func2 = 0


	# If switch variable is 0, the menu will open
	if switch_func3 == 0:

		background.delete('all')
		background.create_image(0, 0, image=background_image, anchor=NW)

		#Defining 2 different search buttons

		payments_button1 = Button(background, text= "Nerasknjižene uplate", fg="black", width=120, height=110, image= unknown_payments_img, compound=TOP, command=unidentified_payments_menu)
		payments_button1_position = background.create_window(322,57, window=payments_button1)
	
		payments_button2 = Button(background, text="Preskočene uplate", fg="black", width=120, height=110, image=skipped_payments_img, compound=TOP, command=skipped_payments)
		payments_button2_position = background.create_window(322,174, window=payments_button2)

		payments_button3 = Button(background, text="Uplate na datum", fg="black", width=120, height=110, image=payments_date_img, compound=TOP)
		payments_button3_position = background.create_window(322,291, window=payments_button3)
		
		switch_func3 = 1
	
	#If the menu is 1 the menu will close	
	else:

		reset_param()


def show_users_by_store():

	'''
	Function used to perform the pre-filtering of users to be displayed inside of the list_box based on the store they belong to.

	1. We set confirm_button1 to be global because we will replicate this whole function for another use, with the difference of confirm_button1. Pressing that butt will activate
		two different functions, so we have to redefine it inside the other function called debt_structure. 

	2. We first define a radiobutton for each store and it's unique value. That value will be assigned to a global variable store_index which is set to "0" by default.
		By pressing a radiobutton we change the store_index value to the value of a specific store. 

	3. Next we define two buttons. One to confirm our choice, and the other one to exit this show_users_by_store menu.

	4. Finally we pre-select the first store as a default choice so that the user is not able to click confirm without at least one option selected.  
	'''

	reset_param()

	global confrim_button1

	button51 = Radiobutton(background, text="MP 51", variable=store_index, value="51", bg="#cbf5e7")
	button51_position = background.create_window(100, 50, window=button51)

	button52 = Radiobutton(background, text="MP 52", variable=store_index, value="52", bg="#cbf5e7")
	button52_position = background.create_window(100, 100, window=button52)

	button53 = Radiobutton(background, text="MP 53", variable=store_index, value="53", bg="#cbf5e7")
	button53_position = background.create_window(100, 150, window=button53)
	
	button54 = Radiobutton(background, text="MP 54", variable=store_index, value="54", bg="#cbf5e7")
	button54_position = background.create_window(100, 200, window=button54)

	button55 = Radiobutton(background, text="MP 55", variable=store_index, value="55", bg="#cbf5e7")
	button55_position = background.create_window(100, 250, window=button55)

	button56 = Radiobutton(background, text="MP 56", variable=store_index, value="56", bg="#cbf5e7")
	button56_position = background.create_window(100, 300, window=button56)

	button59 = Radiobutton(background, text="MP 59", variable=store_index, value="59", bg="#cbf5e7")
	button59_position = background.create_window(100, 350, window=button59)

	button60 = Radiobutton(background, text="MP 60", variable=store_index, value="60", bg="#cbf5e7")
	button60_position = background.create_window(100, 400, window=button60)

	button61 = Radiobutton(background, text="MP 61", variable=store_index, value="61", bg="#cbf5e7")
	button61_position = background.create_window(310, 50, window=button61)						

	button62 = Radiobutton(background, text="MP 62", variable=store_index, value="62", bg="#cbf5e7")
	button62_position = background.create_window(310, 100, window=button62)	

	button63 = Radiobutton(background, text="MP 63", variable=store_index, value="63", bg="#cbf5e7")
	button63_position = background.create_window(310, 150, window=button63)	

	button65 = Radiobutton(background, text="MP 65", variable=store_index, value="65", bg="#cbf5e7")
	button65_position = background.create_window(310, 200, window=button65)	

	button66 = Radiobutton(background, text="MP 66", variable=store_index, value="66", bg="#cbf5e7")
	button66_position = background.create_window(310, 250, window=button66)	

	button67 = Radiobutton(background, text="MP 67", variable=store_index, value="67", bg="#cbf5e7")
	button67_position = background.create_window(310, 300, window=button67)

	button68 = Radiobutton(background, text="MP 68", variable=store_index, value="68", bg="#cbf5e7")
	button68_position = background.create_window(310, 350, window=button68)				

	button69 = Radiobutton(background, text="MP 69", variable=store_index, value="69", bg="#cbf5e7")
	button69_position = background.create_window(310, 400, window=button69)

	button70 = Radiobutton(background, text="MP 70", variable=store_index, value="70", bg="#cbf5e7")
	button70_position = background.create_window(510, 50, window=button70)

	button72 = Radiobutton(background, text="MP 72", variable=store_index, value="72", bg="#cbf5e7")
	button72_position = background.create_window(510, 100, window=button72)

	button73 = Radiobutton(background, text="MP 73", variable=store_index, value="73", bg="#cbf5e7")
	button73_position = background.create_window(510, 150, window=button73)

	button74 = Radiobutton(background, text="MP 74", variable=store_index, value="74", bg="#cbf5e7")
	button74_position = background.create_window(510, 200, window=button74)


	confrim_button1 = Button(background, text="Odaberi" + " "*10, fg="black", width=150, height=30, image=ok_img, compound=RIGHT, command=other_search_criteria)
	confrim_button1_position = background.create_window(98, 520, window=confrim_button1)

	cancel_button1 = Button(background, text="Nazad" + " "*10, fg="black", width=150, height=30, image=cancel_img, compound=RIGHT, command=cancel_by_store)
	cancel_button1_position = background.create_window(542, 520, window=cancel_button1)

	button51.select()

def debt_structure():
	show_users_by_store()

	all_stores = Radiobutton(background, text="Sve MP", variable=store_index, value="Sve MP", bg="#cbf5e7")
	all_stores_position = background.create_window(510, 250, window=all_stores)

	confrim_button1.configure(command=show_debt_structure)


def show_debt_structure():
	reset_param()

	store_stat_data = store_stats_excel(store_index)
	
	credits_total = store_stat_data[0]
	total_value = store_stat_data[1] 
	closed_credits = store_stat_data[2]
	closed_credits_value = store_stat_data[3]
	open_credits = store_stat_data[4]
	open_credits_value = store_stat_data[5]
	late_users = store_stat_data[6]
	late_credits_value = store_stat_data[7]
	total_remaining_value = store_stat_data[8]
	paid_percentage = store_stat_data[11]

	progress = ttk.Progressbar(background, orient = HORIZONTAL, length = 100, mode = 'determinate')
	progress_position = background.create_window(60, 20, window=progress)
	progress['value'] = paid_percentage

	background.create_text(60, 40, text="Isplaćeno " + str(paid_percentage) + "%")


	export_excel_button = Button(background, text=" Eksport tabele    ", fg="black", width=110, height=30, image=excel_img, compound=RIGHT, command=generate_store_workbook)
	export_excel_button_position = background.create_window(574, 25, window=export_excel_button)


	total_value = str(f'{total_value:,}') + " din."
	closed_credits_value = str(f'{closed_credits_value:,}') + " din."
	open_credits_value = str(f'{open_credits_value:,}') + " din."
	total_remaining_value = str(f'{total_remaining_value:,}') + " din."
	late_credits_value = str(f'{late_credits_value:,}') + " din."

	heading_frame = LabelFrame(background)
	heading_frame_position = background.create_window(315,120, window=heading_frame)

	heading_label1 = Label(heading_frame, text="Opis", borderwidth=2, relief='ridge', padx=5, pady=5, width=18, font=("Courier", 14, "bold"))
	heading_label1.grid(row=0, column=0, padx=5, pady=5)

	heading_label2 = Label(heading_frame, text="#", borderwidth=2, relief='ridge', padx=5, pady=5, width=4, font=("Courier", 14, "bold"))
	heading_label2.grid(row=0, column=1, padx=5, pady=5)

	heading_label3 = Label(heading_frame, text="Iznos", borderwidth=2, relief='ridge', padx=5, pady=5, width=18, font=("Courier", 14, "bold"))
	heading_label3.grid(row=0, column=2, padx=5, pady=5)


	row_frame = LabelFrame(background, bg="floralwhite")
	row_frame_position = background.create_window(315,180, window=row_frame)

	row_label1 = Label(row_frame, text="Izdatih kredita", borderwidth=2, relief='ridge', bg="floralwhite", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row_label1.grid(row=0, column=0, padx=5, pady=5)

	row_label2 = Label(row_frame, text=credits_total, borderwidth=2, relief='ridge', bg="floralwhite", padx=5, pady=5, width=5, font=("Courier", 10, "bold"))
	row_label2.grid(row=0, column=1, padx=5, pady=5)

	row_label3 = Label(row_frame, text=total_value, borderwidth=2, relief='ridge', bg="floralwhite", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row_label3.grid(row=0, column=2, padx=5, pady=5)


	row_frame1 = LabelFrame(background, bg="#1abc9c")
	row_frame1_position = background.create_window(315,240, window=row_frame1)

	row1_label1 = Label(row_frame1, text="Isplaćenih kredita", borderwidth=2, relief='ridge', bg="#1abc9c", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row1_label1.grid(row=0, column=0, padx=5, pady=5)

	row1_label2 = Label(row_frame1, text=closed_credits, borderwidth=2, relief='ridge', bg="#1abc9c", padx=5, pady=5, width=5, font=("Courier", 10, "bold"))
	row1_label2.grid(row=0, column=1, padx=5, pady=5)

	row1_label3 = Label(row_frame1, text=closed_credits_value, borderwidth=2, relief='ridge', bg="#1abc9c", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row1_label3.grid(row=0, column=2, padx=5, pady=5)


	row_frame2 = LabelFrame(background, bg="coral1")
	row_frame2_position = background.create_window(315,300, window=row_frame2)

	row2_label1 = Label(row_frame2, text="Aktivnih kredita", borderwidth=2, relief='ridge', bg="coral1", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row2_label1.grid(row=0, column=0, padx=5, pady=5)

	row2_label2 = Label(row_frame2, text=open_credits, borderwidth=2, relief='ridge', bg="coral1",  padx=5, pady=5, width=5, font=("Courier", 10, "bold"))
	row2_label2.grid(row=0, column=1, padx=5, pady=5)

	row2_label3 = Label(row_frame2, text=open_credits_value, borderwidth=2, relief='ridge', bg="coral1",  padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row2_label3.grid(row=0, column=2, padx=5, pady=5)


	row_frame3 = LabelFrame(background, bg="red4")
	row_frame3_position = background.create_window(315,360, window=row_frame3)

	row3_label1 = Label(row_frame3, text="Preminulih korisnika", borderwidth=2, relief='ridge', bg="red4", fg="white", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row3_label1.grid(row=0, column=0, padx=5, pady=5)

	row3_label2 = Label(row_frame3, text=late_users, borderwidth=2, relief='ridge', bg="red4", fg="white", padx=5, pady=5, width=5, font=("Courier", 10, "bold"))
	row3_label2.grid(row=0, column=1, padx=5, pady=5)

	row3_label3 = Label(row_frame3, text=late_credits_value, borderwidth=2, relief='ridge', bg="red4", fg="white", padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row3_label3.grid(row=0, column=2, padx=5, pady=5)

	row_frame4 = LabelFrame(background)
	row_frame4_position = background.create_window(315,420, window=row_frame4)

	row4_label1 = Label(row_frame4, text="Ukupno duga", borderwidth=2, relief='ridge', padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row4_label1.grid(row=0, column=0, padx=(5, 69), pady=5)

	row4_label2 = Label(row_frame4, text=total_remaining_value, borderwidth=2, relief='ridge', padx=5, pady=5, width=25, font=("Courier", 10, "bold"))
	row4_label2.grid(row=0, column=2, padx=5, pady=5)

	confrim_button1 = Button(background, text="Otvori radnju" + " "*4, fg="black", width=150, height=30, image=ok_img, compound=RIGHT, command=other_search_criteria)
	confrim_button1_position = background.create_window(98, 520, window=confrim_button1)

	cancel_button1 = Button(background, text="Nazad" + " "*10, fg="black", width=150, height=30, image=cancel_img, compound=RIGHT, command=debt_structure)
	cancel_button1_position = background.create_window(542, 520, window=cancel_button1)

	if store_index.get() == "Sve MP":
		confrim_button1.configure(state=DISABLED)
		background.create_text(318, 40, text="Sve maloprodaje", font=("Courier", 20, "bold"))

	else:
		background.create_text(318, 25, text="Maloprodaja " + store_index.get(), font=("Courier", 20, "bold"))

def generate_store_workbook():

	generate_store_workbook_excel(store_index)


def sort_by_late_users():

	global store_index

	store_index.set("preminuli")
	other_search_criteria()
	drop.current(0)
	drop_switch(drop.current(0))

def stats_by_store():
	reset_param()

	global drop_menu

	background.create_text(225, 30, text="Statistika po radnjama", font=("Arial", 20, "bold"), tag="title")
	background.create_line(320,60,320,560, width=5)
	background.create_line(0,60,640,60, width=5)
 

	selector_frame = LabelFrame(background)
	selector_frame_position = background.create_window(540, 30, window=selector_frame)

	drop_menu = ttk.Combobox(selector_frame, width= 25, value=["Sortiraj radnje po:", "---------------------------", "Broju izdatih zabrana", "Broju isplaćenih zabrana",
																"Broju preostalih zabrana", "Broju preminulih korisnika", "Vrednosti izdatih zabrana", "Ostatku duga",
																 "Ostatku duga preminulih", "Procentu isplaćenosti", "Prosečnoj vrednosti zabrane", "Najvećoj zabrani"])
	drop_menu.grid(row=0, column=0, padx=5, pady=5)
	drop_menu.current(0)

	drop_menu.bind("<<ComboboxSelected>>", show_store_stats)


def show_store_stats(event):

	background.delete("store_stat", "store_label", "title")
	
	unsorted_data = sort_stores_by_stats_excel(store_index)

	choice = drop_menu.get()

	if choice == "Sortiraj radnje po:" or choice == "---------------------------":
		pass

	else:

		if choice == "Broju izdatih zabrana":
			background.create_text(225, 30, text="Broju izdatih zabrana", font=("Arial", 20, "bold"), tag="title")
			selector = 0
			color = "floralwhite"
			color2 = "black"

		if choice == "Broju isplaćenih zabrana":
			background.create_text(22, 30, text="Broju isplaćenih zabrana", font=("Arial", 20, "bold"), tag="title")
			selector = 2
			color = "#1abc9c"
			color2 = "black"

		if choice == "Broju preostalih zabrana":
			background.create_text(225, 30, text="Broju preostalih zabrana", font=("Arial", 20, "bold"), tag="title")
			selector = 4
			color = "coral1"
			color2 = "black"

		if choice == "Broju preminulih korisnika":
			background.create_text(225, 30, text="Broju preminulih korisnika", font=("Arial", 20, "bold"), tag="title")
			selector = 6
			color = "red4"
			color2 = "white"

		if choice == "Vrednosti izdatih zabrana":
			background.create_text(225, 30, text="Vrednosti izdatih zabrana", font=("Arial", 20, "bold"), tag="title")
			selector = 1
			color = "floralwhite"
			color2 = "black"

		if choice == "Ostatku duga":
			background.create_text(225, 30, text="Ostatku duga", font=("Arial", 20, "bold"), tag="title")
			selector = 8
			color = "coral1"
			color2 = "black"

		if choice == "Ostatku duga preminulih":
			background.create_text(225, 30, text="Ostatku duga preminulih", font=("Arial", 20, "bold"), tag="title")
			selector = 7
			color = "red4"
			color2 = "white"

		if choice == "Procentu isplaćenosti":
			background.create_text(225, 30, text="Procentu isplaćenosti", font=("Arial", 20, "bold"), tag="title")
			selector = 11
			color = "grey"
			color2 = "black"

		if choice == "Prosečnoj vrednosti zabrane":
			background.create_text(225, 30, text="Prosečnoj vrednosti zabrane", font=("Arial", 20, "bold"), tag="title")
			selector = 10
			color = "floralwhite"
			color2 = "black"

		if choice == "Najvećoj zabrani":
			background.create_text(225, 30, text="Najvećoj zabrani", font=("Arial", 20, "bold"), tag="title")
			selector = 9
			color = "floralwhite"
			color2 = "black"


		sorted_list = sorted(unsorted_data, key=operator.itemgetter(selector), reverse=True)

		for item in sorted_list:
			item[1] = str(f'{item[1]:,}') + " din."
			item[8] = str(f'{item[8]:,}') + " din."
			item[7] = str(f'{item[7]:,}') + " din."
			item[10] = str(f'{item[10]:,}') + " din."
			item[9] = str(f'{item[9]:,}') + " din."


		if choice == "Procentu isplaćenosti":

			position_counter = 80
			position_counter2 = 80
			row_counter = 1
			store_counter = 1

			for item in sorted_list:
				if row_counter < 11:
					background.create_text(50, position_counter, text= str(store_counter) + ". MP" + str(item[-1]), tag="store_stat", font=("Arial", 12, "bold"))
					progress = ttk.Progressbar(background, orient = HORIZONTAL, length = 100, mode = 'determinate')
					progress_position = background.create_window(150, position_counter, window=progress, tag="store_label")
					progress['value'] = item[selector]
					background.create_text(220, position_counter, text= str(item[selector]) + "%", tag="store_stat")
					position_counter +=50
					row_counter +=1
					store_counter +=1
				else:
					background.create_text(370, position_counter2, text=str(store_counter) + ". MP" + str(item[-1]), tag="store_stat", font=("Arial", 12, "bold"))
					progress = ttk.Progressbar(background, orient = HORIZONTAL, length = 100, mode = 'determinate')
					progress_position = background.create_window(470, position_counter2, window=progress, tag="store_label")
					progress['value'] = item[selector]
					background.create_text(540, position_counter2, text= str(item[selector]) + "%", tag="store_stat")
					position_counter2 +=50
					row_counter +=1
					store_counter +=1

		else:

			position_counter = 80
			position_counter2 = 80
			row_counter = 1
			store_counter = 1

			for item in sorted_list:
				if row_counter < 11:
					background.create_text(50, position_counter, text= str(store_counter) + ". MP" + str(item[-1]), tag="store_stat", font=("Arial", 12, "bold"))
					store_label = Label(background, text=str(item[selector]), width=20, bg=color, fg=color2, borderwidth=1, relief="solid")
					store_label_position = background.create_window(170, position_counter, window=store_label, tag="store_label")
					position_counter +=50
					row_counter +=1
					store_counter +=1
				else:
					background.create_text(370, position_counter2, text=str(store_counter) + ". MP" + str(item[-1]), tag="store_stat", font=("Arial", 12, "bold"))
					store_label = Label(background, text=str(item[selector]), width=20, bg=color, fg=color2, borderwidth=1, relief="solid")
					store_label_position = background.create_window(490, position_counter2, window=store_label, tag="store_label")
					position_counter2 +=50
					row_counter +=1
					store_counter +=1

def skipped_payments():

	global store_index

	store_index.set("preskočeni")
	other_search_criteria()
	drop.config(value=["Odaberi filter:", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
	drop.current(0)

def populate_entries(user_data):
		

	display_label0.config(text = user_data[0])
	display_label1.config(text = user_data[1])
	display_label2.config(text = user_data[4])
	display_label3.config(text = user_data[3])
	display_label4.config(text = user_data[2])
	display_label5.config(text = f'{user_data[5]:,}' + " din.")
	display_label6.config(text = f'{user_data[10]:,}' + " din.")
	display_label7.config(text = f'{user_data[7]:,}' + " din.")
	display_label8.config(text = user_data[6])
	display_label9.config(text = user_data[9])
	display_label10.config(text = user_data[8])		
	display_label11.config(text = user_data[11])

	if user_data[11] == "Korisnik preminuo":
		display_label11.configure(bg=user_data[12], justify='center', fg="white")
	else:
		display_label11.configure(bg=user_data[12], justify='center', fg="black")

def clear_entries():

	display_label0.config(text = "")
	display_label1.config(text = "")
	display_label2.config(text = "")
	display_label3.config(text = "")
	display_label4.config(text = "")
	display_label5.config(text = "")
	display_label6.config(text = "")
	display_label7.config(text = "")
	display_label8.config(text = "")
	display_label9.config(text = "")
	display_label10.config(text = "")		
	display_label11.config(text = "")
	
	display_label11.configure(bg="white", justify='center')

def drop_switch(event):

	global excel_data
	global index_data
	global user_selected

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120		

	choice = drop.get()
	user_selected = "off"
	clear_entries()
	search_bar.delete(0, END)

	if choice == "Odaberi filter:":
		background.delete("upozorenje")
		left_label.configure(image=question_img_small)
		background.create_text(corx, cory, text="Potrebno je odabrati jedan od ponuđenih filtera", fill="red", tag="upozorenje")

	if choice == "Ime i prezime":
		background.delete("upozorenje")
		left_label.configure(image=name_img_small)

	if choice == "Br. računa":
		background.delete("upozorenje")
		left_label.configure(image=receipt_img_small)

	if choice == "Iznos rate":
		background.delete("upozorenje")
		left_label.configure(image=money_img_small)

	if choice in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]:
		background.delete("upozorenje")
		left_label.configure(image=name_img_small)


	if store_index.get() == "0":
		excel_data, index_data = populate_listbox(drop)


	elif store_index.get() == "preminuli":
		excel_data, index_data = format_by_late_users()

	elif store_index.get() == "preskočeni":
		excel_data, index_data = payments_by_month_excel(drop)

	else:
		excel_data, index_data = format_by_store(store_index)

	list_box.delete(0, END)

	for item in excel_data:
		list_box.insert(END, item)
	
def update_searchbox(excel_data):

	list_box.delete(0, END)

	for item in excel_data:
		list_box.insert(END, item)

def check(e):

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	choice = drop.get()

	if choice == "Odaberi filter:" or choice == None:
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Da biste koristili brzu pretragu, morate prvo odabrati jedan od filtera.", fill="red", tag="upozorenje")

	else:
		background.delete("upozorenje")

		typed = search_bar.get()

		for i,item in enumerate(excel_data):
			if typed.lower() in item.lower():
				list_box.selection_set(i)
			else:
				list_box.selection_clear(i)
		if typed == '':
			list_box.selection_clear(0, END)

def confirm_selection(event):

	global user_value, user_index, user_data, user_instalment_data, user_date_data, user_selected, idx, comment

	w = event.widget

	if len(w.curselection()) == 0:
		pass
	
	else:
		idx = int(w.curselection()[0])
		user_index = index_data[idx]
		user_value = w.get(idx)
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
		background.delete("upozorenje")
		search_bar.delete(0, END)
		search_bar.insert(0, user_value)
		populate_entries(user_data)
		user_selected = "on"

def edit_user():

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120


	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:

		global pop3, pop3_background
		
		pop3 = Toplevel(root)
		pop3.geometry("340x520")		
		pop3.title("PIO EVIDENCIJA 2022")
		pop3.iconbitmap("img_resources/pio.ico")
		pop3.resizable(width=False, height=False)
		pop3.focus_set()                                                        
		pop3.grab_set()

		pop3_background = Canvas(pop3, width=600, height=600)
		pop3_background.pack(side=LEFT, fill=BOTH, expand=1)
		pop3_background_image = ImageTk.PhotoImage(background_img)
		pop3_background.create_image(0, 0, image=background_image, anchor=NW)


		#Defining globals for saving user data in an excel table

		global name2, address2, city2, JMBG2, credit2, instalments2, instalment_info2, receipt2, store2
		
		'''
		Defining a function that takes entries from credit and instalment entries, automaticaly calculates the value of instalment,
		and enters that data into instalment_info entry box
		'''

		def click_on_box2(event):

			credit_data2 = credvar2.get()
				
			if credit_data2 == "":
				pass

			else:
				credit_data2 = credit_data2.replace(",","")
				credit_data2 = float(credit_data2)
				credit2.delete(0, END)
				credit2.insert(0, credit_data2)


		def click_from_box2(event):

			credit_data2 = credvar2.get()
			
			pattern = r"^[1-9]\d*(\.\d+)?$"


				
			if credit_data2 == "":
				pass

			else:
				match = re.match(pattern, credit_data2)
				if match:
					pop3_background.delete("upozorenje")
					credit_data2 = credit_data2.replace(",","")
					credit_data2 = float(credit_data2)
					credit_data2 = str(f'{credit_data2:,}')
					credit2.delete(0, END)
					credit2.insert(0, credit_data2)
				else:
					credit2.delete(0, END)
					pop3_background.delete("upozorenje")
					pop3_background.create_text(165, 460, text="Iznos kredita mora biti upisan u odgovarajućem formatu.", fill="red", tag="upozorenje")	

		def calculate_instalment2(*args):

			credit_data2 = credvar2.get()
			instalment_data2 = instalvar2.get()

			pattern = r"^[1-9]\d*(\.\d+)?$"
			pattern2 = r"^[0-9]*$"

			if instalment_data2 == "" or credit_data2 == "":
				instalment_info2.delete(0, END)

			else:
				credit_data2 = credit_data2.replace(",","")
				match = re.match(pattern, credit_data2)
				match2 = re.match(pattern2, instalment_data2)
				if match and match2:
					pop3_background.delete("upozorenje")
					credit_data2 = float(credit_data2)
					intalment_data2 = int(instalment_data2)
					if int(instalment_data2) > 12 or int(instalment_data2) == 0:
						instalments2.delete(0, END)
						instalment_info2.delete(0, END)
						pop3_background.delete("upozorenje")
						pop3_background.create_text(165, 460, text="Najveći dozvoljeni broj rata je 12.", fill="red", tag="upozorenje")
					else:
						result2 = round(credit_data2/intalment_data2, 2)
						resultvar2.set(result2)
						instalment_info2.delete(0, END)
						instalment_info2.insert(0, str(f'{result2:,}'))
				else:
					instalments2.delete(0, END)
					instalments2.delete(0, END)
					pop3_background.delete("upozorenje")
					pop3_background.create_text(165, 460, text="Kredit i rate moraju biti upisani kao brojevi.", fill="red", tag="upozorenje")		

		
		#Defining description labels

		pop3_background.create_text(70, 50, text="Ime i prezime:")
		pop3_background.create_text(70, 97, text="Adresa:")
		pop3_background.create_text(70, 144, text="Grad:")
		pop3_background.create_text(70, 191, text="JMBG:")
		pop3_background.create_text(70, 238, text="Iznos kredita:")
		pop3_background.create_text(70, 285, text="Broj rata:")
		pop3_background.create_text(70, 332, text="Iznos rate:")
		pop3_background.create_text(70, 379, text="Broj računa:")
		pop3_background.create_text(70, 426, text="Maloprodaja:")	

		#Defining entry boxes for each description label where user data will be entered


		credvar2 = StringVar()
		instalvar2 = StringVar()
		resultvar2 = StringVar()

		name2 = Entry(pop3_background, width=20)
		name2_position = pop3_background.create_window(230, 50, window=name2)

		address2 = Entry(pop3_background, width=20)
		address2_position = pop3_background.create_window(230, 97, window=address2)

		city2 = Entry(pop3_background, width=20)
		city2_position = pop3_background.create_window(230, 144, window=city2)

		JMBG2 = Entry(pop3_background, width=20)
		JMBG2_position = pop3_background.create_window(230, 191, window=JMBG2)

		credit2 = Entry(pop3_background, width=20, textvariable=credvar2)
		credit2_position = pop3_background.create_window(230, 238, window=credit2)

		instalments2 = Entry(pop3_background, width=20, textvariable=instalvar2)
		instalments2_position = pop3_background.create_window(230, 285, window=instalments2)

		instalment_info2 = Entry(pop3_background, width=20, textvariable=resultvar2)
		instalment_info2_position = pop3_background.create_window(230, 332, window=instalment_info2)

		receipt2 = Entry(pop3_background, width=20)
		receipt2_position = pop3_background.create_window(230, 379, window=receipt2)

		store2 = Entry(pop3_background, width=20)
		store2_position = pop3_background.create_window(230, 426, window=store2)

		credit2.delete(0, END)
		instalments2.delete(0, END)
		instalment_info2.delete(0, END)


		credvar2.trace("w", calculate_instalment2)
		instalvar2.trace("w", calculate_instalment2)
		resultvar2.trace("w", calculate_instalment2)

		credit2.bind("<FocusOut>", click_from_box2)
		credit2.bind("<FocusIn>", click_on_box2)


		# Defining two buttons inside of add_user function (Save and Cancel)

		confrim_button2 = Button(pop3_background, text=" Sačuvaj ", fg="black", width=80, height=25, image=ok_img, compound=RIGHT, command=confirm_edit_user)
		confrim_button2_position = pop3_background.create_window(55, 495, window=confrim_button2)

		cancel_button2 = Button(pop3_background, text=" Odustani ", fg="black", width=80, height=25, image=cancel_img, compound=RIGHT, command=close_pop3)
		cancel_button2_position = pop3_background.create_window(285, 495, window=cancel_button2)

		name2.insert(0, user_data[1])
		address2.insert(0, user_data[2])
		city2.insert(0, user_data[3])
		JMBG2.insert(0, user_data[4])
		credit2.insert(0, user_data[5])
		instalments2.insert(0, user_data[6])
		instalment_info2.insert(0, user_data[7])
		receipt2.insert(0, user_data[8])
		store2.insert(0, user_data[9])

		credit_data2 = credvar2.get()
		credit_data2 = credit_data2.replace(",","")
		credit_data2 = float(credit_data2)
		credit_data2 = str(f'{credit_data2:,}')
		credit2.delete(0, END)
		credit2.insert(0, credit_data2)


def confirm_edit_user():

	global user_selected

	if instalment_info2.get() == "":
		pass
	
	else:

		float_instalment_info2 = instalment_info2.get().replace(",","")
		float_instalment_info2 = float(float_instalment_info2)

		float_credit2 = credit2.get().replace(",","")
		float_credit2 = float(float_credit2)

		int_instalments2 = instalments2.get().replace(",","")
		int_instalments2 = int(int_instalments2)

	if name2.get() == "" or address2.get() == "" or city2.get() == "" or JMBG2.get() == "" or credit2.get() == "" or instalments2.get() == "" or instalment_info2.get() == "" or receipt2.get() == "" or store2.get() == "" or float_credit2 < 1 or int_instalments2 < 1 or float_instalment_info2 < 1:
		pop3_background.delete("upozorenje")
		pop3_background.create_text(165, 460, text="Sva polja moraju biti popunjena.", fill="red", tag="upozorenje")

	else:

		try:
			pop3_background.delete("upozorenje")

			data = edit_user_excel(user_index, name2, address2, city2, JMBG2, credit2, instalments2, instalment_info2, receipt2, store2)

			list_box.delete(0, END)
			search_bar.delete(0, END)

			background.create_text(310, 60, text="Promene Sačuvane", fill="red", tag="upozorenje")
			pop3.destroy()
			clear_entries()
			user_selected = "off"

		except ValueError:
				pop3_background.delete("upozorenje")
				pop3_background.create_text(150, 450, text="U poljima iznos kredita, iznos rate i broj rata moraju biti upisani brojevi", fill="red", tag="upozorenje")

def close_pop3():

	name2.delete(0, END)
	address2.delete(0, END)
	city2.delete(0, END)
	JMBG2.delete(0, END)
	credit2.delete(0, END)
	instalments2.delete(0, END)
	instalment_info2.delete(0, END)
	receipt2.delete(0, END)
	store2.delete(0, END)

	pop3.destroy()

def show_full_info():

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:

		global slider_counter, max_slider_value, min_slider_value, left_button, right_button

		slider_counter = index_data.index(user_index)
		max_slider_value = len(index_data) - 1
		min_slider_value = 0

		reset_param()
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)

		name_frame = LabelFrame(background)
		name_frame_position = background.create_window(313,28, window=name_frame)
		name_label = Label(name_frame, text=user_data[1], borderwidth=2, relief='ridge', padx=5, pady=5, font=(16)).pack(padx=5, pady=5)

		status_frame = LabelFrame(background)
		status_frame_position = background.create_window(83, 28, window=status_frame)

		if user_data[11] == "Korisnik preminuo":
			status_label = Label(status_frame, text="Status: " + user_data[11], borderwidth=2, relief='ridge', padx=5, pady=5, width=18, bg="red4", fg="white").pack(padx=5, pady=5)
		
		elif user_data[11] == "Otplata u toku":
			status_label = Label(status_frame, text="Status: " + user_data[11], borderwidth=2, relief='ridge', padx=5, pady=5, width=18, bg="coral1").pack(padx=5, pady=5)

		else:
			status_label = Label(status_frame, text="Status: " + user_data[11], borderwidth=2, relief='ridge', padx=5, pady=5, width=18, bg="#1abc9c").pack(padx=5, pady=5)
			
		store_frame = LabelFrame(background)
		store_frame_position = background.create_window(83, 78, window=store_frame)
		store_label = Label(store_frame, text="Maloprodaja  " + user_data[9], borderwidth=2, relief='ridge', padx=5, pady=5, width=18).pack(padx=5, pady=5)

		credit_frame = LabelFrame(background)
		credit_frame_position = background.create_window(547, 28, window=credit_frame)
		credit_label = Label(credit_frame, text="Početno stanje: " + str(f'{user_data[5]:,}') + " din.", borderwidth=2, relief='ridge', padx=5, pady=5, width=22).pack(padx=5, pady=5)

		balance_frame = LabelFrame(background)
		balance_frame_position = background.create_window(547, 78, window=balance_frame)
		balance_label = Label(balance_frame, text="Ostatak duga: " + str(f'{user_data[10]:,}') + " din.", borderwidth=2, relief='ridge', padx=5, pady=5, width=22).pack(padx=5, pady=5)

		left_button = Button(background, height=50, width=90, image=arrow_left_img, compound=TOP, command=go_left)
		left_button_position = background.create_window(55, 150, window=left_button)

		right_button = Button(background, height=50, width=90, image=arrow_right_img, compound=TOP, command=go_right)
		right_button_position = background.create_window(588, 150, window=right_button)

		col_heading_frame = LabelFrame(background)
		col_heading_frame_position = background.create_window(320, 240, window=col_heading_frame)

		instalments_frame = LabelFrame(background)
		instalments_frame_position = background.create_window(320, 280, window=instalments_frame)

		dates_frame = LabelFrame(background)
		dates_frame_position = background.create_window(320, 320, window=dates_frame)


		bottom_frame = LabelFrame(background, width=80)
		bottom_frame_position = background.create_window(319, 380, window=bottom_frame)



		payment_button = Button(bottom_frame, text="Uplata", image=pay_img_small, compound=TOP, width=98, command=payment_full_info_menu)
		payment_button.grid(row=0, column=0, padx=6, pady=5)			
		
		payment2_button = Button(bottom_frame, text="Neplanska uplata", image=pay2_img_small, compound=TOP, width=99, command=atypical_payment_menu)
		payment2_button.grid(row=0, column=1, padx=6, pady=5)

		cancel_payment_button = Button(bottom_frame, text="Obriši uplatu", image=cancel_pay_img_small, compound=TOP, width=99, command=cancel_payment_full_info_menu)
		cancel_payment_button.grid(row=0, column=2, padx=7, pady=5)

		late_user_button = Button(bottom_frame, text="Korisnik preminuo", image=cross_img_small, compound=TOP, width=99, command=late_user)
		late_user_button.grid(row=0, column=3, padx=7, pady=5)

		print_data_button = Button(bottom_frame, text="Štampa", image=print_img_small, compound=TOP, width=98, command=print_user)
		print_data_button.grid(row=0, column=4, padx=7, pady=5)

		confrim_button = Button(background, text="Napomene" + " "*10, fg="black", width=150, height=30, image=note_img, compound=RIGHT, command=add_comment)
		confrim_button_position = background.create_window(98, 520, window=confrim_button)

		cancel_button = Button(background, text="Nazad" + " "*10, fg="black", width=150, height=30, image=cancel_img, compound=RIGHT, command=close_full_info)
		cancel_button_position = background.create_window(542, 520, window=cancel_button)

		if slider_counter >= max_slider_value:
			right_button.config(state=DISABLED)

		if slider_counter == 0:
			left_button.config(state=DISABLED)



		possible_table_rows = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
		t_rows = []

		for no in range(user_data[6]):
			t_rows.append(possible_table_rows[no])

		row_counter2 = 0
		for item in t_rows:
			Label(col_heading_frame, text=item, borderwidth=2, relief='ridge', height=2, width=2, padx=16).grid(row=0,column=row_counter2)
			row_counter2 +=1

		if user_data[11] == "Korisnik preminuo":

			row_counter3 = 0
			for item in user_instalment_data:
				Label(instalments_frame, text=item, borderwidth=2, relief='ridge', height=2,  width=2, padx=16, bg="red4", fg="white").grid(row=0,column=row_counter3)
				row_counter3 +=1

			row_counter4 = 0
			for date in user_date_data:
				Label(dates_frame, text=date, borderwidth=2, relief='ridge', height=2, width=2, padx=16,  bg="red4", fg="white").grid(row=0,column=row_counter4)
				row_counter4 +=1

		elif user_data[11] == "Kredit otplaćen":

			row_counter3 = 0
			for item in user_instalment_data:
				Label(instalments_frame, text=item, borderwidth=2, relief='ridge', height=2,  width=2, padx=16, bg="#1abc9c").grid(row=0,column=row_counter3)
				row_counter3 +=1

			row_counter4 = 0
			for date in user_date_data:
				Label(dates_frame, text=date, borderwidth=2, relief='ridge', height=2, width=2, padx=16,  bg="#1abc9c").grid(row=0,column=row_counter4)
				row_counter4 +=1


		else:
			row_counter3 = 0
			for item in user_instalment_data:
				if item == 0:
					Label(instalments_frame, text=str(f'{item:,}'), borderwidth=2, relief='ridge', height=2,  width=2, padx=16, bg="coral1").grid(row=0,column=row_counter3)
					row_counter3 +=1
				else:
					Label(instalments_frame, text=str(f'{item:,}'), borderwidth=2, relief='ridge', height=2,  width=2, padx=16, bg="#1abc9c").grid(row=0,column=row_counter3)
					row_counter3 +=1

			row_counter4 = 0
			for date in user_date_data:
				if date == None:
					Label(dates_frame, text=date, borderwidth=2, relief='ridge', height=2, width=2, padx=16,  bg="coral1").grid(row=0,column=row_counter4)
					row_counter4 +=1
				else:
					Label(dates_frame, text=date, borderwidth=2, relief='ridge', height=2, width=2, padx=16,  bg="#1abc9c").grid(row=0,column=row_counter4)
					row_counter4 +=1

def go_right():

	global slider_counter, user_selected, user_index, idx, user_data, user_instalment_data, user_date_data, comment

	slider_counter +=1

	if slider_counter <= max_slider_value:
		user_index = index_data[slider_counter]
		idx = slider_counter
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
		user_selected = "on"
		show_full_info()
		search_bar.delete(0, END)
		search_bar.insert(0, user_data[1])
	else:
		right_button.config(state=DISABLED)

def go_left():

	global slider_counter, user_selected, user_index, idx, user_data, user_instalment_data, user_date_data, comment

	slider_counter -= 1

	if slider_counter >= min_slider_value:
		user_index = index_data[slider_counter]
		idx = slider_counter
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
		user_selected = "on"
		show_full_info()
		search_bar.delete(0, END)
		search_bar.insert(0, user_data[1])
	else:
		left_button.config(state=DISABLED)


def add_comment():

	print("selecting func")

	global pop9, pop9_entry

	pop9 = Toplevel(root)
	pop9.geometry("600x250")		
	pop9.title("PIO EVIDENCIJA 2022")
	pop9.iconbitmap("img_resources/pio.ico")
	pop9.resizable(width=False, height=False)
	pop9.focus_set()                                                        
	pop9.grab_set()


	pop9_background = Canvas(pop9, width=300, height=175)
	pop9_background.pack(side=LEFT, fill=BOTH, expand=1)
	pop9_background_image = ImageTk.PhotoImage(background_img)
	pop9_background.create_image(0, 0, image=background_image, anchor=NW)

	pop9_label = Label(pop9_background, text="Napomene: ", borderwidth=2, relief='ridge', padx=5, pady=5, width=80)
	pop9_label_position = pop9_background.create_window(300, 18, window=pop9_label)

	pop9_entry = Text(pop9_background, width=71, height=10)
	pop9_entry_position = pop9_background.create_window(300, 120, window=pop9_entry)
	pop9_entry.insert(1.0, comment)

	pop9_confrim_button = Button(pop9_background, text="Sačuvaj ", fg="black", width=80, height=25, image=ok_img, compound=RIGHT, command=write_comment)
	pop9_confrim_button_position = pop9_background.create_window(50, 228, window=pop9_confrim_button)

	pop9_no_button = Button(pop9_background, text="Odustani ", fg="black", width=80, height=25, image=cancel_img, compound=RIGHT, command=close_pop9)
	pop9_no_button_position = pop9_background.create_window(550, 228, window=pop9_no_button)


def close_pop9():

	pop9.destroy()

def write_comment():

	global comment, user_selected

	user_selected = "on"

	user_comment = pop9_entry.get(1.0, END)
	write_comment_excel(user_comment, user_index)
	pop9.destroy()

	show_full_info()
	user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
	


def close_full_info():

	'''
	Function used to get back to main menu that is other_search_criteria.

	1. First we check whether we have to get back to main menu with all users or main menu that has been filtered by store. So we check the store_index.get().
	If this is true we set the drop.current to 0 instead of 1 because the main menu by store has been stripped of all other drop menu options except the "Ime i prezime" option.

	2. Everything else work exactly the same and we take the following steps. 
		a. We save the name from seach_bar to re-insert it.
		b. We call other_search criteria to get back to that menu.
		c. We call the drop.get() to check what was the selected option to re-use it again.
		d. We use the re-selected drop value inside the drop_switch func to re-populate the list_box with users. 
		e. We call the global idx which acts as the index of specific user inside of the listbox. 
			Note that this idx is different from user_index which corresponds to the index in the excel table
			That way we can automaticaly re-select the user in the list_box when returning to the main menu
		f. We re-insert the name in search_bar that we saved in the step a.
		g. We populate the display labels by calling the global user_data that we obtained back in confirm_selection func
		h. Finally we set the global user_selected to "on" to be able to preform the main menu button actions on re-selected user. This option was set to off inside of the drop_switch func.

	'''

	global user_selected

	if store_index.get() != "0":

		

		if store_index.get() == "preskočeni":
			if menu_view == "wide":
				other_search_criteria()
				user_selected = "off"

			else:
				other_search_criteria()
				other_search_criteria_narrow()
				user_selected = "off"
		
		elif store_index.get() == "preminuli":
			if menu_view == "wide":
				other_search_criteria()
				user_selected = "off"

			else:
				other_search_criteria()
				other_search_criteria_narrow()
				user_selected = "off"

		else:
			if menu_view == "wide":
				name = search_bar.get()
				other_search_criteria()
				drop.current(0)
				drop_switch(drop.current(0))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
				user_selected = "on"
			else:
				name = search_bar.get()
				other_search_criteria()
				other_search_criteria_narrow()
				drop.current(0)
				drop_switch(drop.current(0))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
				user_selected = "on"

	else:

		if drop.get() == "Ime i prezime":
			if menu_view == "wide":
				name = search_bar.get()
				other_search_criteria()
				drop.current(1)
				drop_switch(drop.current(1))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
			else:
				name = search_bar.get()
				other_search_criteria()
				other_search_criteria_narrow()
				drop.current(1)
				drop_switch(drop.current(1))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
		
		if drop.get() == "Br. računa":
			if menu_view == "wide":
				name = search_bar.get()
				other_search_criteria()
				drop.current(2)
				drop_switch(drop.current(2))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
			else:
				name = search_bar.get()
				other_search_criteria()
				other_search_criteria_narrow()
				drop.current(2)
				drop_switch(drop.current(2))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
		
		if drop.get() == "Iznos rate":
			if menu_view == "wide":
				name = search_bar.get()
				other_search_criteria()
				drop.current(3)
				drop_switch(drop.current(3))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)
			else:
				name = search_bar.get()
				other_search_criteria()
				other_search_criteria_narrow()
				drop.current(3)
				drop_switch(drop.current(3))
				list_box.select_set(idx)
				search_bar.insert(0, name)
				populate_entries(user_data)

		user_selected = "on"
	
def payment():

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:
	
		global user_data	

		if user_data[11] == "Korisnik preminuo":
			background.delete("upozorenje")
			background.create_text(corx, cory, text="Nije moguće proknjižiti uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")

		elif user_data[11] == "Kredit otplaćen":
			background.delete("upozorenje")
			background.create_text(corx, cory, text=f"Kredit za korisnika {user_data[1]} je već isplaćen.", fill="red", tag="upozorenje")


		else:
			background.delete("upozorenje")
			instalment_num = payment_excel(user_index, today)
			background.create_text(corx, cory, text= f"Proknjižena rata br. {instalment_num} za korisnika {user_data[1]}", fill="red", tag="upozorenje")

			
			user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
			populate_entries(user_data)

def cancel_payment():

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:

		global user_data

		if user_data[11] == "Korisnik preminuo":
			background.delete("upozorenje")
			background.create_text(corx, cory, text="Nije moguće obrisati uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")			

		else:
			background.delete("upozorenje")
			instalment_num = cancel_payment_excel(user_index)
			background.create_text(corx, cory, text=f"Jedna rata obrisana za korisnika {user_data[1]}. Trenutno proknjiženo {instalment_num} od {user_data[6]} rata", fill="red", tag="upozorenje")

			user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)
			populate_entries(user_data)

def payment_full_info_menu():

	global user_selected, user_data

	user_selected = "on"


	if user_data[11] == "Korisnik preminuo":
		background.delete("upozorenje")
		background.create_text(315, 150, text="Nije moguće proknjižiti uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")


	elif user_data[11] == "Kredit otplaćen":
			background.delete("upozorenje")
			background.create_text(315, 150, text=f"Kredit za korisnika {user_data[1]} je već isplaćen.", fill="red", tag="upozorenje")

	else:
		background.delete("upozorenje")
		payment_excel(user_index, today)
		show_full_info()
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)



def cancel_payment_full_info_menu():

	global user_selected, user_data

	user_selected = "on"

	if user_data[11] == "Korisnik preminuo":
		background.delete("upozorenje")
		background.create_text(315, 150, text="Nije moguće obrisati uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")

	else:
		background.delete("upozorenje")
		cancel_payment_excel(user_index)
		show_full_info()
		user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)

def print_user():

	print_user_excel(user_index)

def atypical_payment_menu():

	global pop, pop_background
	global date_entry
	global instalment_entry

	pop = Toplevel(root)
	pop.geometry("200x160")		
	pop.title("PIO EVIDENCIJA 2022")
	pop.iconbitmap("img_resources/pio.ico")
	pop.resizable(width=False, height=False)
	pop.focus_set()                                                        
	pop.grab_set()
 

	pop_background = Canvas(pop, width=400, height=250)
	pop_background.pack(side=LEFT, fill=BOTH, expand=1)
	pop_background_image = ImageTk.PhotoImage(background_img)
	pop_background.create_image(0, 0, image=background_image, anchor=NW)

	pop_background.create_text(76, 23, text="Datum uplate:")
	pop_background.create_text(53, 73, text="Iznos:")

	
	
	date_entry = Entry(pop_background, width=20)
	date_entry_position = pop_background.create_window(100, 38, window=date_entry)
	date_entry.insert(0, today)

	instalment_entry = Entry(pop_background, width=20)
	instalment_entry_position = pop_background.create_window(100, 88, window=instalment_entry)
	instalment_entry.insert(0, user_data[7])

	pop_confrim_button = Button(pop_background, text="Ok ", fg="black", width=48, height=25, image=ok_img, compound=RIGHT, command=atypical_payment)
	pop_confrim_button_position = pop_background.create_window(32, 138, window=pop_confrim_button)

	pop_cancel_button = Button(pop_background, text="Ne ", fg="black", width=48, height=25, image=cancel_img, compound=RIGHT, command=close_popup)
	pop_cancel_button_position = pop_background.create_window(168, 138, window=pop_cancel_button)

def atypical_payment():

	pattern = r"^([0-2][0-9]|(3)[0-1])(\.)(((0)[0-9])|((1)[0-2]))(\.)\d{2}$"

	pattern2 = r"^[1-9]\d*(\.\d+)?$"

	date = date_entry.get()
	instalment = instalment_entry.get()

	match = re.match(pattern, date)
	match2 = re.match(pattern2, instalment)

	global user_selected, user_data

	user_selected = "on"

	pop_background.delete("upozorenje")


	if user_data[11] == "Korisnik preminuo":
		background.delete("upozorenje")
		background.create_text(315, 150, text="Nije moguće proknjižiti uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")

	elif user_data[11] == "Kredit otplaćen":
		background.delete("upozorenje")
		background.create_text(315, 150, text=f"Kredit za korisnika {user_data[1]} je već isplaćen.", fill="red", tag="upozorenje")

	else:

		if instalment_entry.get() == "" or date_entry.get() == "":
			pop_background.delete("upozorenje")
			pop_background.create_text(100, 107, text="Sva polja moraju biti popunjena", fill="red", tag="upozorenje")

		elif not match:
			pop_background.delete("upozorenje")
			pop_background.create_text(100, 107, text="Forma datuma mora biti dd.mm.yy", fill="red", tag="upozorenje")

		elif not match2:
			pop_background.delete("upozorenje")
			pop_background.create_text(100, 107, text="U polju iznos mora biti upisan broj", fill="red", tag="upozorenje")

		else:
			background.delete("upozorenje")	
			atypical_payment_excel(user_index, instalment_entry, date_entry)
			show_full_info()
			user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)

		'''
		else:

			try:
				background.delete("upozorenje")	
				atypical_payment_excel(user_index, instalment_entry, date_entry)
				show_full_info()
				user_data, user_instalment_data, user_date_data = confirm_selection_excel(user_index)
			
			except ValueError:
				pop_background.delete("upozorenje")
				pop_background.create_text(100, 107, text="U polju iznos mora biti upisan broj", fill="red", tag="upozorenje")
		'''

def close_popup():
	pop.destroy()

def late_user():

	global user_selected, user_data

	user_selected = "on"

	late_user_excel(user_index, today)
	show_full_info()

	user_data, user_instalment_data, user_date_data, comment = confirm_selection_excel(user_index)

	'''
	if status_color == "red4":
		pygame.mixer.music.load("Hallelujah.mp3")
		pygame.mixer.music.play(loops=0)

	else:
		pygame.mixer.music.load("death.mp3")
		pygame.mixer.music.play(loops=0)
	'''	


def delete_user_popup():

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	global user_selected

	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:
		global pop2 
		pop2 = Toplevel(root)
		pop2.geometry("200x160")		
		pop2.title("PIO EVIDENCIJA 2022")
		pop2.iconbitmap("img_resources/pio.ico")
		pop2.resizable(width=False, height=False)
		pop2.focus_set()                                                        
		pop2.grab_set()


		pop2_background = Canvas(pop2, width=200, height=175)
		pop2_background.pack(side=LEFT, fill=BOTH, expand=1)
		pop2_background_image = ImageTk.PhotoImage(background_img)
		pop2_background.create_image(0, 0, image=background_image, anchor=NW)

		
		pop2_background.create_image(77, 5, image=warning_img, anchor=NW)
		

		pop2_background.create_text(100, 66, text= "Da li želite da obrišete korisnika:", fill="red")
		pop2_background.create_text(100, 66, text= "_________________________________", fill="red")

		pop2_background.create_text(105, 99, text= user_data[1])


		pop2_confrim_button = Button(pop2_background, text=" Da ", fg="black", width=48, height=25, image=ok_img, compound=RIGHT, command=delete_user)
		pop2_confrim_button_position = pop2_background.create_window(32, 138, window=pop2_confrim_button)

		pop2_cancel_button = Button(pop2_background, text=" Ne ", fg="black", width=48, height=25, image=cancel_img, compound=RIGHT, command=close_popup2)
		pop2_cancel_button_position = pop2_background.create_window(168, 138, window=pop2_cancel_button)


def delete_user():

	global user_selected

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120


	def delete_warning2():
		background.delete("upozorenje")

	
	background.delete("upozorenje")
	delete_user_excel(user_index)
	pop2.destroy()
	background.create_text(corx, cory, text="Korisnik " + user_data[1] + " obrisan iz baze korisnika.", fill="red", tag="upozorenje")
	list_box.delete(0, END)
	search_bar.delete(0, END)
	drop.current(0)
	clear_entries()
	background.after(5000, delete_warning2)
	user_selected = "off"


def close_popup2():
	pop2.destroy()

def serial_payments():

	global list_box, index_data

	if menu_view == "wide":
		corx = 310
		cory = 60
	else:
		corx = 420
		cory = 120

	if user_selected == "off":
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Prvo odaberite filter, pa korisnika, pa neku od opcija", fill="red", tag="upozorenje")
	
	else:

		search_list = list_box.curselection()

		user_index_list = []

		for item in search_list:
			user_index_list.append(index_data[item])

		no_of_customers = serial_payments_excel(user_index_list, today)
		background.delete("upozorenje")
		background.create_text(corx, cory, text="Proknjižene uplate za " + str(no_of_customers) + " korisnika",fill="red", tag="upozorenje")	
	

def other_search_criteria():
	reset_param()

	global menu_view

	menu_view = "wide"

	global drop, search_bar, list_box, left_label, left_frame, middle_frame, bottom_frame
	global left_frame_position, middle_frame_position, box_frame_position, data_frame_position, bottom_frame_position
	global edit_button, delete_button, payment_button, serial_payment_button, cancel_payment_button, enlarge_button, view_button
	global label0, label1, label2, label3, label4, label5, label6, label7, label8, label9, label10, label11
	global display_label0, display_label1, display_label2, display_label3, display_label4, display_label5, display_label6, display_label7, display_label8, display_label9, display_label10, display_label11

	left_frame = LabelFrame(background)
	left_frame_position = background.create_window(76, 25, window=left_frame)

	middle_frame = LabelFrame(background)
	middle_frame_position = background.create_window(320, 25, window=middle_frame)

	box_frame = LabelFrame(background, width=80)
	box_frame_position = background.create_window(323, 162, window=box_frame)

	data_frame = LabelFrame(background, width=80, padx=7)
	data_frame_position = background.create_window(323, 367, window=data_frame)

	bottom_frame = LabelFrame(background, width=80)
	bottom_frame_position = background.create_window(320, 521, window=bottom_frame)

	drop = ttk.Combobox(left_frame, width= 12, value=["Odaberi filter:", "Ime i prezime", "Br. računa", "Iznos rate"])
	drop.grid(row=0, column=0, padx=5, pady=5)
	drop.current(0)
	drop.bind("<<ComboboxSelected>>", drop_switch)

	left_label = Label(left_frame, image=question_img_small)
	left_label.grid(row=0, column=1)

	middle_label = Label(middle_frame, image=search_img_small)
	middle_label.grid(row=0, column=1)

	stats_button_shortcut = Button(background, text="Struktura duga   ", width=136, image=calc_small_img, compound=RIGHT, command=show_debt_structure)
	stats_button_shortcut_postition = background.create_window(564, 25, window=stats_button_shortcut)

	view_button = Button(background, width=45, image=narrow_view_img, command=other_search_criteria_narrow)
	view_button_postition = background.create_window(35, 250, window=view_button)

	search_bar = Entry(middle_frame, width=25)
	search_bar.grid(row=0, column=2, padx=5, pady=9)

	list_box = Listbox(box_frame, width=80, height=10, selectbackground='#1abc9c', selectmode=EXTENDED)
	list_box.grid(row=0, column=0, columnspan=4, padx=10, pady=(8,0))

	list_box.bind('<<ListboxSelect>>', confirm_selection)
	search_bar.bind("<KeyRelease>", check)

	label0 = Label(data_frame, text="Indeks:")
	label0.grid(row=1, column=0, padx=10, pady=(4,1), sticky=W)

	display_label0 = Label(data_frame, bg="white", width=2, relief="sunken", bd=1)
	display_label0.grid(row=1, column=0, padx=(60,0), pady=(4,1), sticky=W)

	label1 = Label(data_frame, text="Ime i Prezime")
	label1.grid(row=2, column=0, padx=10, pady=10, sticky=W)

	display_label1 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label1.grid(row=2, column=1, padx=10, pady=6)

	label2 = Label(data_frame, text="JMBG")
	label2.grid(row=2, column=2, padx=10, pady=6, sticky=W)

	display_label2 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label2.grid(row=2, column=3, padx=10, pady=6)

	label3 = Label(data_frame, text="Grad")
	label3.grid(row=3, column=0, padx=10, pady=6, sticky=W)

	display_label3 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label3.grid(row=3, column=1, padx=10, pady=6)

	label4 = Label(data_frame, text="Adresa")
	label4.grid(row=3, column=2, padx=10, pady=6, sticky=W)

	display_label4 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label4.grid(row=3, column=3, padx=10, pady=6)

	label5 = Label(data_frame, text="Iznos kredita")
	label5.grid(row=4, column=0, padx=10, pady=6, sticky=W)

	display_label5 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label5.grid(row=4, column=1, padx=10, pady=6)

	label6 = Label(data_frame, text="Ostatak duga")
	label6.grid(row=4, column=2, padx=10, pady=6, sticky=W)

	display_label6 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label6.grid(row=4, column=3, padx=10, pady=6)

	label7 = Label(data_frame, text="Iznos rate")
	label7.grid(row=5, column=0, padx=10, pady=6, sticky=W)

	display_label7 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label7.grid(row=5, column=1, padx=10, pady=6)

	label8 = Label(data_frame, text="Broj rata")
	label8.grid(row=5, column=2, padx=10, pady=6, sticky=W)

	display_label8 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label8.grid(row=5, column=3, padx=10, pady=6)

	label9 = Label(data_frame, text="Maloprodaja")
	label9.grid(row=6, column=0, padx=10, pady=6, sticky=W)

	display_label9 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label9.grid(row=6, column=1, padx=10, pady=6)

	label10 = Label(data_frame, text="Broj računa")
	label10.grid(row=6, column=2, padx=10, pady=6, sticky=W)

	display_label10 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	display_label10.grid(row=6, column=3, padx=10, pady=6)

	label11 = Label(data_frame, text="Trenutni status", padx=10, pady=6)
	label11.grid(row=7, column=0)

	display_label11 = Label(data_frame, bg="white", width=52, relief="sunken", bd=1)
	display_label11.grid(row=7, column=1, columnspan=3, padx=10, pady=6)


	edit_button = Button(bottom_frame, text="Izmeni", image=edit_img_small, compound=TOP, width=80, command=edit_user)
	edit_button.grid(row=0, column=0, padx=3, pady=5)

	delete_button = Button(bottom_frame, text="Obriši korisnika", image=trash_img_small, compound=TOP, width=80, command=delete_user_popup)
	delete_button.grid(row=0, column=1, padx=3, pady=5)

	payment_button = Button(bottom_frame, text="Uplata", image=pay_img_small, compound=TOP, width=80, command=payment)
	payment_button.grid(row=0, column=2, padx=3, pady=5)

	serial_payment_button = Button(bottom_frame, text="Serijska uplata", image=money_img_small, compound=TOP, width=80, command=serial_payments)
	serial_payment_button.grid(row=0, column=3, padx=3, pady=5)

	cancel_payment_button = Button(bottom_frame, text="Obriši uplatu", image=cancel_pay_img_small, compound=TOP, width=80, command=cancel_payment)
	cancel_payment_button.grid(row=0, column=4, padx=3, pady=5)

	enlarge_button = Button(bottom_frame, text="Pun prikaz", image=enlarge_img_small, compound=TOP, width=80, command=show_full_info)
	enlarge_button.grid(row=0, column=5, padx=3, pady=5)

	if store_index.get() != "0":

		if store_index.get() == "preskočeni":
			drop.config(value=["Odaberi filter:", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
			drop.current(0)

		else:

			drop.config(value=["Ime i prezime"])
			drop.current(0)
			drop_switch(drop.current(0))
	
	if store_index.get() == "preminuli" or store_index.get() == "0" or store_index.get() == "preskočeni":
		stats_button_shortcut.config(state=DISABLED)

def other_search_criteria_narrow():

	global menu_view, view_button2

	menu_view = "narrow"

	view_button.destroy()

	background.coords("upozorenje", 420, 120)

	background.coords(left_frame_position, 564, 70)

	background.coords(middle_frame_position, 350, 25)
	
	background.coords(box_frame_position, 100, 274)
	list_box.config(height=33, width=30)
	list_box.grid(row=0, column=0, padx=1, pady=1)

	
	bottom_frame.config(width=20)
	background.coords(bottom_frame_position, 570, 341)

	background.coords(data_frame_position, 360, 340)

	label0.grid(row=0, column=0, pady=(9,2))
	display_label0.grid(row=0, column=1, padx=(10,0), pady=(9,2))
	label1.grid(row=1, column=0)
	display_label1.grid(row=1, column=1)
	label2.grid(row=2, column=0)
	display_label2.grid(row=2, column=1)
	label3.grid(row=3, column=0)
	display_label3.grid(row=3, column=1)
	label4.grid(row=4, column=0)
	display_label4.grid(row=4, column=1)
	label5.grid(row=5, column=0)
	display_label5.grid(row=5, column=1)
	label6.grid(row=6, column=0)
	display_label6.grid(row=6, column=1)
	label7.grid(row=7, column=0)
	display_label7.grid(row=7, column=1)
	label8.grid(row=8, column=0)
	display_label8.grid(row=8, column=1)
	label9.grid(row=9, column=0)
	display_label9.grid(row=9, column=1)
	label10.grid(row=10, column=0)
	display_label10.grid(row=10, column=1)
	label11.grid(row=11, column=0)
	display_label11.grid(row=11, column=1)
	display_label11.config(width=17)
	

	edit_button.grid(row=0, column=0, padx=3, pady=5)
	delete_button.grid(row=1, column=0, padx=3, pady=5)
	payment_button.grid(row=2, column=0, padx=3, pady=5)
	serial_payment_button.grid(row=3, column=0, padx=3, pady=5)
	cancel_payment_button.grid(row=4, column=0, padx=3, pady=5)
	enlarge_button.grid(row=5, column=0, padx=3, pady=5)

	view_button2 = Button(background, width=136, height=30, image=wide_view_img, compound=RIGHT, command=return_to_wide_view)
	view_button_postition2 = background.create_window(350, 70, window=view_button2)

def return_to_wide_view():

	global menu_view, view_button

	menu_view = "wide"

	view_button2.destroy()

	background.coords("upozorenje", 310, 60)

	background.coords(left_frame_position, 76, 25)

	background.coords(middle_frame_position, 320, 25)
	
	background.coords(box_frame_position, 323, 162)
	list_box.config(height=10, width=80)
	list_box.grid(row=0, column=0, columnspan=4, padx=10, pady=(8,0))

	
	bottom_frame.config(width=80)
	background.coords(bottom_frame_position, 320, 521)

	background.coords(data_frame_position, 323, 367)

	label0.grid(row=1, column=0, padx=10, pady=(4,1))
	display_label0.grid(row=1, column=0, padx=(60,0), pady=(4,1))

	label1.grid(row=2, column=0)
	display_label1.grid(row=2, column=1)

	label2.grid(row=2, column=2)
	display_label2.grid(row=2, column=3)

	label3.grid(row=3, column=0)
	display_label3.grid(row=3, column=1)

	label4.grid(row=3, column=2)
	display_label4.grid(row=3, column=3)

	label5.grid(row=4, column=0)
	display_label5.grid(row=4, column=1)

	label6.grid(row=4, column=2)
	display_label6.grid(row=4, column=3)

	label7.grid(row=5, column=0)
	display_label7.grid(row=5, column=1)

	label8.grid(row=5, column=2)
	display_label8.grid(row=5, column=3)

	label9.grid(row=6, column=0)
	display_label9.grid(row=6, column=1)

	label10.grid(row=6, column=2)
	display_label10.grid(row=6, column=3)

	label11.grid(row=7, column=0)
	display_label11.grid(row=7, column=1, columnspan=3, padx=10, pady=6)

	display_label11.config(width=52)
	

	edit_button.grid(row=0, column=0, padx=3, pady=5)
	delete_button.grid(row=0, column=1, padx=3, pady=5)
	payment_button.grid(row=0, column=2, padx=3, pady=5)
	serial_payment_button.grid(row=0, column=3, padx=3, pady=5)
	cancel_payment_button.grid(row=0, column=4, padx=3, pady=5)
	enlarge_button.grid(row=0, column=5, padx=3, pady=5)

	view_button = Button(background, width=45, image=narrow_view_img, command=other_search_criteria_narrow)
	view_button_postition = background.create_window(35, 250, window=view_button)



def confirm_by_store():
	
	reset_param()
	other_search_criteria()
	drop.config(value=["Ime i prezime"])
	drop.current(0)
	drop_switch(drop.current(0))
	
def cancel_by_store():
	reset_param()

def populate_unidentified_entries():
	 label02.config(text =unidentified_user_data[0])
	 label04.config(text = f'{unidentified_user_data[1]:,}' + " din.")
	 label06.config(text = unidentified_user_data[2])

def onselect(event):
	global value
	global unidentified_user_data
	global uni_select

	w = event.widget
	uni_select = 1

	if len(w.curselection()) == 0:
		pass
	
	else:
		idx = int(w.curselection()[0])
		value = w.get(idx)
		unidentified_user_data = select_unidentified_user_excel(value)
		populate_unidentified_entries()
		

def unidentified_payments_menu():
	reset_param()

	global unidentified_search_box, label02, label04, label06, uni_select

	uni_select = 0

	data = populate_unidentified_search_box()

	data_frame = LabelFrame(background, width=80)
	data_frame_position = background.create_window(323, 260, window=data_frame)

	bottom_frame = LabelFrame(background, width=80)
	bottom_frame_position = background.create_window(320, 520, window=bottom_frame)

	unidentified_search_box = Listbox(data_frame, width=37, height=20, selectbackground='#1abc9c', selectmode=SINGLE)
	unidentified_search_box.grid(row=0, column=0, columnspan=4, padx=5, pady=5)

	label01 = Label(data_frame, text="Ime i prezime:")
	label01.grid(row=2, column=0, padx=10, pady=6, sticky=W)

	label02 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	label02.grid(row=2, column=1, padx=10, pady=6)

	label03 = Label(data_frame, text="Iznos uplate")
	label03.grid(row=3, column=0, padx=10, pady=6, sticky=W)

	label04 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	label04.grid(row=3, column=1, padx=10, pady=6)

	label05 = Label(data_frame, text="Datum uplate")
	label05.grid(row=4, column=0, padx=10, pady=6, sticky=W)

	label06 = Label(data_frame, bg="white", width=17, relief="sunken", bd=1)
	label06.grid(row=4, column=1, padx=10, pady=6)

	add_payment_button = Button(bottom_frame, text="Dodaj uplatu", image=add_payment_img, compound=TOP, width=85, command=add_payment_pop)
	add_payment_button.grid(row=0, column=0, padx=3, pady=5)

	delete_payment_button = Button(bottom_frame, text="Obriši uplatu", image=trash_img_small, compound=TOP, width=85, command=delete_payment_pop)
	delete_payment_button.grid(row=0, column=1, padx=3, pady=5)

	connect_payment_button = Button(bottom_frame, text="Proknjiži uplatu", image=pay_img_small, compound=TOP, width=85, command=connect_payment_step1)
	connect_payment_button.grid(row=0, column=2, padx=3, pady=5)

	for item in data:
		unidentified_search_box.insert(END, item)

	unidentified_search_box.bind('<<ListboxSelect>>', onselect)

def add_payment_pop():

	global pop4, unidentified_payment, unidentified_name, pop4_background

	pop4 = Toplevel(root)
	pop4.geometry("200x160")		
	pop4.title("PIO EVIDENCIJA 2022")
	pop4.iconbitmap("img_resources/pio.ico")
	pop4.resizable(width=False, height=False)
	pop4.focus_set()                                                        
	pop4.grab_set()   


	pop4_background = Canvas(pop4, width=400, height=250)
	pop4_background.pack(side=LEFT, fill=BOTH, expand=1)
	pop4_background_image = ImageTk.PhotoImage(background_img)
	pop4_background.create_image(0, 0, image=background_image, anchor=NW)

	pop4_background.create_text(76, 23, text="Ime i Prezime:")
	pop4_background.create_text(53, 73, text="Iznos:")


	unidentified_name = Entry(pop4_background, width=20)
	unidentified_name_position = pop4_background.create_window(100, 38, window=unidentified_name)

	unidentified_payment = Entry(pop4_background, width=20)
	unidentified_payment_position = pop4_background.create_window(100, 88, window=unidentified_payment)

	pop4_confrim_button = Button(pop4_background, text=" Da ", fg="black", width=48, height=25, image=ok_img, compound=RIGHT, command=add_payment)
	pop4_confrim_button_position = pop4_background.create_window(32, 138, window=pop4_confrim_button)

	pop4_cancel_button = Button(pop4_background, text=" Ne ", fg="black", width=48, height=25, image=cancel_img, compound=RIGHT, command=close_pop4)
	pop4_cancel_button_position = pop4_background.create_window(168, 138, window=pop4_cancel_button)
	
def add_payment():

	if unidentified_payment.get() == "" or unidentified_name.get() == "":
		pop4_background.delete("upozorenje")
		pop4_background.create_text(100, 107, text="Sva polja moraju biti popunjena", fill="red", tag="upozorenje")
	
	else:
		try:
			pop4_background.delete("upozorenje")
			add_payment_excel(unidentified_payment, unidentified_name, today)
			unidentified_payment.delete(0, END)
			unidentified_name.delete(0, END)
			pop4.destroy()
			unidentified_payments_menu()
		except ValueError:
			pop4_background.delete("upozorenje")
			pop4_background.create_text(100, 107, text="U polju iznos mora biti upisan broj", fill="red", tag="upozorenje")


def delete_payment_pop():

	if uni_select == 0:
		background.delete("upozorenje")
		background.create_text(324, 20, text="Prvo odaberite korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:
		background.delete("upozorenje")

		global pop5 
		pop5 = Toplevel(root)
		pop5.geometry("200x160")		
		pop5.title("PIO EVIDENCIJA 2022")
		pop5.iconbitmap("img_resources/pio.ico")
		pop5.resizable(width=False, height=False)
		pop5.focus_set()                                                        
		pop5.grab_set()


		pop5_background = Canvas(pop5, width=200, height=175)
		pop5_background.pack(side=LEFT, fill=BOTH, expand=1)
		pop5_background_image = ImageTk.PhotoImage(background_img)
		pop5_background.create_image(0, 0, image=background_image, anchor=NW)

		
		pop5_background.create_image(77, 5, image=warning_img, anchor=NW)
		pop5_background.create_text(100, 66, text= "Da li želite da obrišete uplatu?")


		pop5_confrim_button = Button(pop5_background, text=" Da ", fg="black", width=48, height=25, image=ok_img, compound=RIGHT, command=delete_payment)
		pop5_confrim_button_position = pop5_background.create_window(32, 138, window=pop5_confrim_button)

		pop5_cancel_button = Button(pop5_background, text=" Ne ", fg="black", width=48, height=25, image=cancel_img, compound=RIGHT, command=close_pop5)
		pop5_cancel_button_position = pop5_background.create_window(168, 138, window=pop5_cancel_button)


def delete_payment():
	delete_unidentified_payment(value)
	unidentified_payments_menu()
	pop5.destroy()

def close_pop4():

	pop4.destroy()

def close_pop5():

	pop5.destroy()

def several_users_pop():

	global pop6, pop6_background, uni_index
	pop6 = Toplevel(root)
	pop6.geometry("300x200")		
	pop6.title("PIO EVIDENCIJA 2022")
	pop6.iconbitmap("img_resources/pio.ico")
	pop6.resizable(width=False, height=False)
	pop6.focus_set()                                                        
	pop6.grab_set()


	pop6_background = Canvas(pop6, width=200, height=175)
	pop6_background.pack(side=LEFT, fill=BOTH, expand=1)
	pop6_background_image = ImageTk.PhotoImage(background_img)
	pop6_background.create_image(0, 0, image=background_image, anchor=NW)

	pop6_background.create_text(145, 20, text= f"Ima više korisnika pod imenom {uni_name}" , fill="red")

	pop6_background.create_text(110, 60, text= f"U ovo polje unestie\njednistveni indeks korisnika")

	pop6_background.create_text(92, 100, text= f"Moguće opcije: {uni_indexes}")

	uni_index = Entry(pop6_background, width=5)
	uni_index_position = pop6_background.create_window(250, 60, window=uni_index)

	help_button = Button(pop6_background, width=40, height=25, image=question_img, command=payment_help)
	help_button_position = pop6_background.create_window(250, 95, window=help_button)


	pop6_confrim_button = Button(pop6_background, text=" Proknjiži ", fg="black", width=88, height=25, image=ok_img, compound=RIGHT, command=connect_payment_confirm)
	pop6_confrim_button_position = pop6_background.create_window(55, 178, window=pop6_confrim_button)

	pop6_cancel_button = Button(pop6_background, text=" Odustani ", fg="black", width=88, height=25, image=cancel_img, compound=RIGHT, command=close_pop6)
	pop6_cancel_button_position = pop6_background.create_window(245, 178, window=pop6_cancel_button)

def close_pop6():

	pop6.destroy()

def payment_help():
	global pop7
	pop7 = Toplevel(root)
	pop7.geometry("481x508")		
	pop7.title("PIO EVIDENCIJA 2022")
	pop7.iconbitmap("img_resources/pio.ico")
	pop7.resizable(width=False, height=False)
	pop7.focus_set()                                                        
	pop7.grab_set()

	pop7_background = Canvas(pop7)
	pop7_background.pack(side=LEFT, fill=BOTH, expand=1)
	pop7_background.create_image(0, 0, image=help_img, anchor=NW)

	pop7.protocol('WM_DELETE_WINDOW', on_exit)

def on_exit():

	pop7.destroy()
	pop6.focus_set()                                                        
	pop6.grab_set()

def connect_payment_step1():


	if uni_select == 0:
		background.delete("upozorenje")
		background.create_text(324, 20, text="Prvo odaberite korisnika, pa neku od opcija", fill="red", tag="upozorenje")

	else:
		background.delete("upozorenje")

		global uni_name, uni_indexes, unidentified_data

		unidentified_data = select_unidentified_user_excel(value)

		uni_name = unidentified_data[0]

		
		condtition, uni_indexes = initial_unidentified_search(uni_name)

		if condtition == 0:
			background.delete("upozorenje")
			background.create_text(324, 20, text=f"Korisnik pod imenom {uni_name} ne postoji u bazi korisnika", fill="red", tag="upozorenje")


		elif condtition == 1:
			background.delete("upozorenje")
			data, instalment_data, date_data = find_unidentified_user(uni_name)
			
			if data[11] == "Korisnik preminuo":
				background.delete("upozorenje")
				background.create_text(324, 20, text="Nije moguće proknjižiti uplatu za korisnika koji je preminuo", fill="red", tag="upozorenje")

			elif data[11] == "Kredit otplaćen":
				background.delete("upozorenje")
				background.create_text(324, 20, text=f"Kredit za korisnika {data[1]} je već isplaćen.", fill="red", tag="upozorenje")

			else:
				background.delete("upozorenje")
				atypical_payment_excel(data[0], unidentified_data[1], unidentified_data[2])
				delete_unidentified_payment(value)
				unidentified_payments_menu()
				background.create_text(324, 20, text=f"Uplata proknjižena na korisnika {data[1]}.", fill="red", tag="upozorenje")
		
		else:
			several_users_pop()
			

def connect_payment_confirm():

	print(type(uni_index.get()))


	if uni_index.get() == "":
		pop6_background.delete("upozorenje")
		pop6_background.create_text(145, 140, text="Unesite jedinstveni indeks", fill="red", tag="upozorenje")

	else:
		try:

			if int(uni_index.get()) not in uni_indexes:
				pop6_background.delete("upozorenje")
				pop6_background.create_text(145, 140, text=f"Uneti indeks ne odgovara korisniku {uni_name}", fill="red", tag="upozorenje")
			
			else:
				pop6_background.delete("upozorenje")
				data, instalment_data, date_data, comment = confirm_selection_excel(int(uni_index.get()))

				if data[11] == "Korisnik preminuo":
					pop6_background.delete("upozorenje")
					pop6_background.create_text(145, 140, text="Nije moguće proknjižiti uplatu za korisnika\n\t   koji je preminuo", fill="red", tag="upozorenje")

				elif data[11] == "Kredit otplaćen":
					pop6_background.delete("upozorenje")
					pop6_background.create_text(145, 140, text=f"Kredit za korisnika {data[1]} je već isplaćen.", fill="red", tag="upozorenje")

				else:
					atypical_payment_excel(int(uni_index.get()), unidentified_data[1], unidentified_data[2])
					close_pop6()
					delete_unidentified_payment(value)
					unidentified_payments_menu()
					background.create_text(324, 20, text=f"Uplata proknjižena na korisnika {data[1]}.", fill="red", tag="upozorenje") 
		except:
			pop6_background.delete("upozorenje")
			pop6_background.create_text(145, 140, text="U polju indeks mora biti upisan broj", fill="red", tag="upozorenje")

def exit_program():

	global pop8, pop8_background, exit_image

	if login_var == "on":

		pop8 = Toplevel(root)
		pop8.geometry("300x160")		
		pop8.title("PIO EVIDENCIJA 2022")
		pop8.iconbitmap("img_resources/pio.ico")
		pop8.resizable(width=False, height=False)
		pop8.focus_set()                                                        
		pop8.grab_set()

		pop8_background = Canvas(pop8, width=300, height=175)
		pop8_background.pack(side=LEFT, fill=BOTH, expand=1)

		exit_img = Image.open("img_resources/background.jpg")
		exit_img = exit_img.resize((330, 200), Image.ANTIALIAS)
		exit_image = ImageTk.PhotoImage(exit_img)

		pop8_background.create_image(0, 0, image=exit_image, anchor=NW)
		pop8_background.create_image(130, 5, image=warning_img, anchor=NW)
		

		pop8_background.create_text(155, 66, text= "Da li želite da sačuvate promene?", fill="red")
		pop8_background.create_text(155, 66, text= "___________________________________", fill="red")


		pop8_confrim_button = Button(pop8_background, text=" Da     ", fg="black", width=80, height=25, image=ok_img, compound=RIGHT, command=close_program)
		pop8_confrim_button_position = pop8_background.create_window(50, 138, window=pop8_confrim_button)

		pop8_no_button = Button(pop8_background, text=" Ne     ", fg="black", width=80, height=25, image=cancel_img, compound=RIGHT, command=close_without_saving)
		pop8_no_button_position = pop8_background.create_window(150, 138, window=pop8_no_button)

		pop8_cancel_button = Button(pop8_background, text=" Odustani ", fg="black", width=80, height=25, image=return_img, compound=RIGHT, command=cancel_close)
		pop8_cancel_button_position = pop8_background.create_window(250, 138, window=pop8_cancel_button)

	else:
		root.destroy()

def close_program():

	save_document_excel()
	root.destroy()


def cancel_close():
	pop8.destroy()

def close_without_saving():
	root.destroy()

#create_excel_sheet()
#create_unidentified_payments_excel_table()

login_page()

initial_load()



root.protocol('WM_DELETE_WINDOW', exit_program)

root.mainloop()