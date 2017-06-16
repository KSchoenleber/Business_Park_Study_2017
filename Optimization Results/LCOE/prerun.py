import os
import tkinter.font as tkFont
import sys
import math
import matplotlib.pyplot as plt
import pandas as pd
import pyomo.environ
import pyomo.core as pyomo
import shutil
import xlrd
from tkinter import *
from math import *
from tkinter import messagebox
from openpyxl import load_workbook
from xlrd import XLRDError
from datetime import datetime

class application:
	def __init__(self, master):
		self.master = master
		self.frame = Frame(self.master)
		self.master.title("Evaluate LCOE")
		# width x height + x_offset + y_offset:
		# self.master.geometry("500x500")
		self.master.resizable(0, 0)
		self.customFont = tkFont.Font(family="Arial", size=9)
		self.input_widgets()
		self.buttons()
		self.directory_path()

	def input_widgets(self):
		# input filename
		Label(self.master, text="Inputfile").grid(sticky=W, column=0, row=0, padx=5, pady=5)
		self.filename = Entry(self.master, width=40, font=self.customFont)
		self.filename.grid(column=1, row=0, padx=5, pady=5)
		self.filename_quote = """mimo-example.xlsx"""
		self.filename.insert(END, self.filename_quote)
		# input processes
		Label(self.master, text="Processes").grid(sticky=W, column=0, row=2, padx=5, pady=5)
		self.pro = Text(self.master, height=4, width=40, font=self.customFont)
		self.pro.grid(column=1, row=2, padx=5, pady=5)
		self.pro_quote = """Hydro plant, Wind park, Photovoltaics, Gas plant"""
		self.pro.insert(END, self.pro_quote)
		# input process chain
		Label(self.master, text="Process chain").grid(sticky=W, column=0, row=3, padx=5, pady=5)
		self.proch = Text(self.master, height=2, width=40, font=self.customFont)
		self.proch.grid(column=1, row=3, padx=5, pady=5)
		
	def buttons(self):
		self.start = Button(self.master, text='Start evaluation', command=self.input_data)
		self.start.grid(row=7, column=0,ipadx=5, padx=5, pady=5)
		self.quit = Button(self.master, text='Quit', command=self.master.quit)
		self.quit.grid(sticky=W, row=7, column=1,ipadx=5, padx=5, pady=5)
		
	def directory_path(self):
		# return path of result directory
		self.path = Label(self.master, anchor=W, justify=LEFT)
		self.path.grid(sticky=W, column=0, columnspan=2, row=5, padx=5, pady=5)
		
	def input_data(self):	
		self.input_file = str(self.filename.get())
		input_pros = str(self.pro.get('1.0', 'end-1c'))
		pchain = str(self.proch.get('1.0', 'end-1c'))	
		if not os.path.exists(self.input_file):
			messagebox.showwarning("Open file", "Cannot open this input file\n{}".format(self.input_file))
		else:
			data(self.input_file, input_pros, pchain)
			result_path = data.prepare_result_directory(self, self.input_file)
			self.path.configure(text="Result directory:\n{}".format(result_path))

class data:
	def __init__(self, input_file, input_pros, pchain):
	
		self.input_file = input_file
		self.input_pros = input_pros
		self.pchain = pchain
		
		with pd.ExcelFile(self.input_file) as xls:
			try:
				site = xls.parse('Site').set_index(['Name'])
				commodity = (
					xls.parse('Commodity').set_index(['Site', 'Commodity']))
				process = xls.parse('Process').set_index(['Site', 'Process'])
				self.process_commodity = (
					xls.parse('Process-Commodity')
					   .set_index(['Process', 'Commodity', 'Direction']))
				transmission = (
					xls.parse('Transmission')
					   .set_index(['Site In', 'Site Out',
								   'Transmission', 'Commodity']))
				storage = (
					xls.parse('Storage').set_index(['Site', 'Storage', 'Commodity']))
				demand = xls.parse('Demand').set_index(['t'])
				supim = xls.parse('SupIm').set_index(['t'])
				buy_sell_price = xls.parse('Buy-Sell-Price').set_index(['t'])
				dsm = xls.parse('DSM').set_index(['Site', 'Commodity'])
			except XLRDError:
				sys.exit("One or more main sheets are missing in your Input-file. Compare it with mimo-example.xlsx!")
			try:
				hacks = xls.parse('Hacks').set_index(['Name'])
			except XLRDError:
				hacks = None

		# prepare input data
		# split columns by dots '.', so that 'DE.Elec' becomes the two-level
		# column index ('DE', 'Elec')
		demand.columns = self.split_columns(demand.columns, '.')
		supim.columns = self.split_columns(supim.columns, '.')
		buy_sell_price.columns = self.split_columns(buy_sell_price.columns, '.')

		data = {
			'site': site,
			'commodity': commodity,
			'process': process,
			'process_commodity': self.process_commodity,
			'transmission': transmission,
			'storage': storage,
			'demand': demand,
			'supim': supim,
			'buy_sell_price': buy_sell_price,
			'dsm': dsm}
		if hacks is not None:
			data['hacks'] = hacks

		# sort nested indexes to make direct assignments work
		for key in data:
			if isinstance(data[key].index, pd.core.index.MultiIndex):
				data[key].sortlevel(inplace=True)


		# get used sites and "main" commoditys
		msites = list(demand.columns.levels[0])
		mcom = list(demand.columns.levels[1])

		# get required supim data
		self.msupim = supim[msites].sum().to_frame().rename(columns={0:'FLH'})

		# get required demand data
		mdemand = demand[msites].sum().to_frame().rename(columns={0:'Demand'})

		# get required process data 
		mpara = ['inv-cost', 'fix-cost', 'var-cost', 'wacc', 'depreciation']
		self.mprocess = process.loc[msites, mpara]

		# get commodity type - SupIm, Stock
		com_type = commodity.loc[msites, 'Type'].to_frame().reset_index(['Site', 'Commodity'])
		com_type = com_type[(com_type['Type'] == "SupIm") | (com_type['Type'] == "Stock")].fillna(0)
		com_type = com_type.drop(labels=['Site'], axis = 1)
		self.com_type = com_type.drop_duplicates(['Commodity'], keep='last').set_index('Commodity')
		com = list(self.com_type.index)
		com2 = mcom + com

		# get required process-commodity data
		ratio_in = self.process_commodity.loc[(slice(None), com, 'In'), :]['ratio'].to_frame()
		self.ratio_out = self.process_commodity.loc[(slice(None), com2, 'Out'), :]['ratio'].to_frame()
		
		# get processes from input
		pros = list(ratio_in.index.levels[0])
		self.input_pros = self.input_pros.split(', ')

		# get related commodity for each process
		self.pro_com = ratio_in.reset_index(['Direction', 'Commodity']).drop(['Direction', 'ratio'], axis = 1)

		# process chain inputs and intermediate commodity
		self.pchain = self.pchain.split(', ')
		if self.pchain == ['']: x = 0
		else: x = len(self.pchain)

		# get CO2 ratios
		self.ratio_outco2 = self.process_commodity.xs(('CO2', 'Out'), level=['Commodity','Direction'])['ratio'].to_frame()

		# get commodity price
		self.com_price = commodity.loc[msites, 'price'].to_frame().rename(columns={0:'price'}).fillna(0)

		# calculation for commoditys with electricity output
		for site in msites:
			# reset process chain counter for next site
			p = 0
			result_site = pd.DataFrame()
			result_site_suf = pd.DataFrame()    
			
			for mc in mcom:
				# initialize result Dataframe
				LCOE_con = pd.DataFrame()
				LCOE_reg = []
				
				for pro in self.input_pros:
					if pro not in pros:
						print("{} is not a valid Process! Valid inputs are: {}".format(pro, pros))
						break
					
					elif pro in (self.mprocess.ix[site].index.get_level_values(0)):     
						if mc in self.process_commodity.ix[pro].index.get_level_values(0):
							self.calc_result(site, pro, LCOE_reg, LCOE_con, mc)
							if (p < (x-2)) & (pro in self.pchain): p += 1
						
						elif (pro in self.pchain[p:x]) & (self.pchain != ['']):   
							if (pro in self.pchain[p:x]) & (mc in self.process_commodity.ix[self.pchain[p+1]].index.get_level_values(0)):
								self.calc_result(site, pro, LCOE_reg, LCOE_con, self.pro_com.get_value(self.pchain[p+1], 'Commodity'))                
								if p < (x-2): p += 1
								
						else:
							continue
				
				if len(LCOE_reg) != 0:
					LCOE_reg = pd.DataFrame(LCOE_reg).set_index('FLH')
				else:
					LCOE_reg = pd.DataFrame()
					
				if len(LCOE_con) != 0:
					LCOE_con.index.name = 'FLH'
				
				# combine results (renewable, conventional processes) of current site 
				result_mc = pd.concat([LCOE_con, LCOE_reg], axis=1)
				result_mc_suf = result_mc.add_suffix('_{}'.format(mc))
				result_site = pd.concat([result_site, result_mc], axis=1)
				result_site_suf = pd.concat([result_site_suf, result_mc_suf], axis=1).fillna(0)
				
				# plot each site as png-file
				if not result_mc.empty:
					self.plot_site(LCOE_reg, LCOE_con, result_mc, result_site, mc, site)
			
			# export result of each site to LCOE.xlsx
			self.result_export(result_site_suf, site)
	
	def split_columns(self, columns, sep='.'):
		"""Split columns by separator into MultiIndex.

		Given a list of column labels containing a separator string (default: '.'),
		derive a MulitIndex that is split at the separator string.

		Args:
			columns: list of column labels, containing the separator string
			sep: the separator string (default: '.')

		Returns:
			a MultiIndex corresponding to input, with levels split at separator

		Example:
			>>> split_columns(['DE.Elec', 'MA.Elec', 'NO.Wind'])
			MultiIndex(levels=[['DE', 'MA', 'NO'], ['Elec', 'Wind']],
					   labels=[[0, 1, 2], [0, 0, 1]])

		"""
		if len(columns) == 0:
			return columns
		column_tuples = [tuple(col.split('.')) for col in columns]
		return pd.MultiIndex.from_tuples(column_tuples)

	def calc_annuity(self, r, n):
		q = 1 + r
		a = ((q ** n) * (q - 1)) / ((q ** n) - 1)
		return a

	def calc_LCOE(self, invc, fixc, varc, fuelc, co2c, eff, total_eff, FLH, r, n):
		# calculate annuity investment costs
		invc_a = invc * self.calc_annuity(r, n)
		# calculate LOCE
		LCOE = (((invc_a + fixc) / FLH) + varc + ((fuelc + co2c) / eff))# * (eff / total_eff)
		return LCOE

	def calc_result(self, site, pro, LCOE_reg, LCOE_con, mc):
		
		icost = self.mprocess.loc[(site, pro), 'inv-cost']
		fixcost = self.mprocess.loc[(site, pro), 'fix-cost']
		varcost = self.mprocess.loc[(site, pro), 'var-cost']
		wacc = self.mprocess.loc[(site, pro), 'wacc']
		dep = self.mprocess.loc[(site, pro), 'depreciation']
		eff = self.process_commodity.loc[(pro, mc, 'Out'),'ratio']
		total_eff = self.process_commodity.loc[pro]['ratio'].sum()
		
		pcom = self.pro_com.loc[pro, 'Commodity']
		fuelcost = self.com_price.loc[(site, pcom), 'price']
		if (pro in self.pchain) & (fuelcost == 0):
			try: fuelcost = LCOE_reg.loc[:, self.pchain[self.pchain.index(pro)-1]] 
			except:
				try: fuelcost = LCOE_con.loc[:, self.pchain[self.pchain.index(pro)-1]]
				except: 
					try: fuelcost = result_site.loc[:, self.pchain[self.pchain.index(pro)-1]]
					except: fuelcost = 0
		
		if pro in self.ratio_outco2.index:
			co2cost = self.ratio_outco2.loc[pro, 'ratio']
		else:
			co2cost = 0

		if (self.com_type.loc[pcom, 'Type'] == "SupIm"):
			FLH = self.msupim.loc[(site, pcom), 'FLH'].astype(int)
			
			result_reg = self.calc_LCOE(icost, fixcost, varcost, fuelcost, co2cost, eff, total_eff, FLH, wacc, dep)
			LCOE_reg.append({pro: result_reg, 'FLH': FLH})
			return LCOE_reg

		else:
			FLH = range(1,8761)
			
			result_con = self.calc_LCOE(icost, fixcost, varcost, fuelcost, co2cost, eff, total_eff, FLH, wacc, dep)
			result_con = pd.DataFrame({pro: result_con})
			LCOE_con.insert(0, pro, result_con)
			return LCOE_con

	def prepare_result_directory(self, input_file):
		""" create a time stamped directory within the result folder """
		# timestamp for result directory
		now = datetime.now().strftime('%Y%m%dT%H%M')

		# create result directory if not existent
		result_name = os.path.splitext(input_file)[0]  # cut away file extension
		result_dir = os.path.join('evaluate_LCOE', '{}-{}'.format(result_name, now))
		if not os.path.exists(result_dir):
			os.makedirs(result_dir)

		return result_dir

	def plot_site(self, LCOE_reg, LCOE_con, result_mc, result_site, mc , site):
			# plot figure of current site
			
			# initialize x-axis and legend
			hour_con = list(range(1, 8761))
			hour_reg = list(LCOE_reg.index)
			legend = list(result_mc.columns)
			
			# plot reg and con LCOE in one plot
			if len(LCOE_con.columns) != 0:
				plt.plot(hour_con, LCOE_con)
			
			if len(LCOE_reg.columns) != 0:
				plt.plot(hour_reg, LCOE_reg, 'o')
			
			plt.title('LCOE - {}'.format(mc))
			plt.xlabel('hour')
			plt.ylabel('LCOE [€/MWh]')
			plt.xlim([1,8760])
			plt.ylim([0,(result_site.min().max()*3)])
			plt.rcParams['legend.numpoints'] = 1
			plt.legend(legend, fontsize=8)
			
			# save plot of each side to specific path        
			result_dir = self.prepare_result_directory(self.input_file)  # name + time stamp
			plt.savefig('{}\LCOE_{}_{}.png'.format(result_dir, mc, site), dpi=400)
			plt.close()

	def result_export(self, file, site):
		
		result_dir = self.prepare_result_directory(self.input_file)
		
		if os.path.exists('{}\LCOE.xlsx'.format(result_dir, site)):
			with pd.ExcelWriter('{}\LCOE.xlsx'.format(result_dir, site), engine='openpyxl') as writer:
				writer.book = load_workbook('{}\LCOE.xlsx'.format(result_dir, site))
				file.to_excel(writer, sheet_name=site)
		else:
			writer = pd.ExcelWriter('{}\LCOE.xlsx'.format(result_dir, site), engine='xlsxwriter')
			file.to_excel(writer, sheet_name=site)
			writer.save()

	def multiindex(self, columns, mc):
		columns = list(columns)
		column_tuples = [(mc, col) for col in columns]
		return pd.MultiIndex.from_tuples(column_tuples)	
	
def main(): 
    root = Tk()
    app = application(root)
    root.mainloop()

if __name__ == '__main__':
    main()

