from selenium import webdriver
import re
import xlsxwriter
import sys

class TelegramMessage:	

	def __init__(self, HTML_element):
		self.date = TelegramMessage.determine_date(HTML_element)
		self.sender = TelegramMessage.determine_sender(HTML_element)
		self.text = TelegramMessage.determine_text(HTML_element)

	def determine_date(HTML_element):
		try:
			return HTML_element.find_element_by_xpath(".//div[contains(@class, 'date details')]").get_attribute("title")
		except Exception:
			return 0

	def determine_sender(HTML_element):
		try:
			return HTML_element.find_element_by_xpath(".//div[@class='from_name']").text
		except Exception:
			return "--joined"

	def determine_text(HTML_element):
		try:
			return HTML_element.find_element_by_xpath(".//div[@class='text']").text
		except Exception:
			return "[Image]"

#получение URL'а из консоли
url = sys.argv[1]

#вызов браузера
browser = webdriver.Chrome("include\\chromedriver.exe")
browser.minimize_window()
browser.get(url)

html_messages = browser.find_elements_by_xpath("//div[contains(@class, 'message default')]")

#конвертация сообщений в массив
messages_array = []
for msg in html_messages:
	messages_array.append(TelegramMessage(msg))

#обработка --joined
assembled_index = 0
excess = []
for i in range(len(messages_array)):
	if (messages_array[i].sender != "--joined"):
		if ((messages_array[i].sender == messages_array[assembled_index].sender) 
			and ("Смотрим" in messages_array[assembled_index].text)):
				messages_array[assembled_index].text += "\n\n" + messages_array[i].text
		assembled_index = i
	else:
		if ("✅" in messages_array[assembled_index].text):
			messages_array[assembled_index].text = messages_array[i].text
			excess.append(i)
		else:
			messages_array[assembled_index].text += "\n\n" + messages_array[i].text
			excess.append(i)

for index in reversed(excess):
	messages_array.pop(index)

#создание таблицы в нужном формате
messages_table = []
for i in range(len(messages_array)):
	if (i < (len(messages_array) - 1)):
		bot_re = re.compile("MosruQaBot|Mos.ru")
		if (bot_re.match(messages_array[i].sender)) and not (bot_re.match(messages_array[i + 1].sender)) and ("Выключенные сборки:" not in messages_array[i].text):
			messages_table.append([messages_array[i].date, messages_array[i + 1].sender, messages_array[i].text, " ", " ", " ", " ", " ", " ", messages_array[i + 1].text])
			ticket_re = re.compile("[A-Z]+-\d{1,5}")
			hpsm_re = re.compile("[A-Z]{1,2}\d{8}")
			if (ticket_re.search(messages_array[i + 1].text)):
				messages_table[-1][7] = "https://olymp-moscow.atlassian.net/browse/" + ticket_re.search(messages_array[i + 1].text)[0]
			elif (hpsm_re.search(messages_array[i + 1].text)):
				messages_table[-1][7] = hpsm_re.search(messages_array[i + 1].text)[0]


#экспорт в xlsx
workbook = xlsxwriter.Workbook(f"report.xlsx")
worksheet = workbook.add_worksheet()
for i in range(len(messages_table)):
	for j in range(len(messages_table[i])):
		worksheet.write(i, j, messages_table[i][j])
workbook.close()

print("That's All, Folks!")
browser.quit()