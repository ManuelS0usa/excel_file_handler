import os
import xlrd


class Excel:

	def __init__( self, someFile ):

		self.fileToOpen = someFile
		self.fileContent = self.readFile()
		self.allSheetNames = self.getSheetNames()
		self.allMaxLins = [ self.getMaxLinesBySheetIndex(self.allSheetNames.index(l)) for l in self.allSheetNames ]
		self.allMaxCols = [ self.getMaxColsBySheetIndex(self.allSheetNames.index(c)) for c in self.allSheetNames ]
		#self.fileName = getFileName()

	def _closesession( self ):
		
		try:
			self.__session.close()
		except:
			pass

	def readFile( self ):

		''' Read excel file passed as parameter. 
			Returns list of dicts w/ file content, each dict representing one worksheet as follows:
			[
				{ 	'folha': workSheetName, 'linhas': nLines, 'colunas': nRows, 
					'dados': [ 
						[ { 'tipo': cellType, 'valor': cellValue }, { other cell... } ], 
						[ { other line... } ]
					]
				}
			]
			Note: Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
		'''
		workBook   = xlrd.open_workbook( self.fileToOpen )
		allSheetNames = workBook.sheet_names()
		livroExcel = []

		for sheetNum in range( 0, len(allSheetNames) ):
			dadosList = []
			workSheet = workBook.sheet_by_index( sheetNum )

			nlinhas = workSheet.nrows
			ncolunas = workSheet.ncols
			
			num_rows = nlinhas - 1
			num_cells = ncolunas - 1

			curr_row = -1
			while curr_row < num_rows:
				curr_row += 1
				row = workSheet.row(curr_row)
				curr_cell = -1
				linhaTemp = []
				while curr_cell < num_cells:
					curr_cell += 1
					cell_type = workSheet.cell_type(curr_row, curr_cell)
					cell_value = workSheet.cell_value(curr_row, curr_cell)
					linhaTemp.append( {'tipo':cell_type, 'valor':cell_value} )
				dadosList.append( linhaTemp )

			livroExcel.append({'folha':workSheet.name, 'linhas':nlinhas, 'colunas':ncolunas,'dados':dadosList})
		
		return livroExcel

	def getFileName( self ):

		return self.fileToOpen.split('/')[-1]

	def getSheetNames( self ):

		return [ f['folha'] for f in self.fileContent ]

	def getMaxLins( self ):

		return self.allMaxLins

	def getMaxCols( self ):

		return self.allMaxCols

	# -----------
	def getDataBySheetName( self, sName ):

		return [ f['dados'] for f in self.fileContent if f['folha'] == sName ][0]

	def getMaxLinesBySheetName( self, sName ):
		
		return [ f['linhas'] for f in self.fileContent if f['folha'] == sName ][0]

	def getMaxRowsBySheetName( self, sName ):
		
		return [ f['colunas'] for f in self.fileContent if f['folha'] == sName ][0]
	# -----------

	def getDataBySheetIndex( self, whatIndex ):

		return self.fileContent[whatIndex]['dados']

	def getMaxLinesBySheetIndex( self, whatIndex ):
		
		return self.fileContent[whatIndex]['linhas']

	def getMaxColsBySheetIndex( self, whatIndex ):
		
		return self.fileContent[whatIndex]['colunas']

	def parseExcelDate( self, dateNum ):

		''' Convert excel week day reference to textual data (pt) '''

		pt  = { 1:'Dom', 2:'Seg', 3:'Ter', 4:'Qua', 5:'Qui', 6:'Sex', 7:'Sab' }
		# day = time.strftime( "%A", time.strptime( dateNum, "%Y%m%d" ) )

		try:
			if dateNum > 0 and dateNum <= 7:
				dia = [ v for k,v in pt.items() if k == dateNum ][0]
			else:
				rawDate = ''
				rawDate = xlrd.xldate_as_tuple( dateNum, 0 )[0:3]			
				dia = '-'.join( str(i) for i in rawDate )
			return dia

		except:
			return dateNum

	def select( self, sheetIndex, whatCels ):

		''' Returns a list of lists (one per line) w/ selected file content.
			Parameters:
				sheetIndex = integer number of the desired sheet (first is 0)
				whatCels = [ (ini line , ini column), (end line, end column) ]. «None» in ini values => 0 | «None» in end values => all
		'''

		selectedData = []
		allData = self.getDataBySheetIndex( sheetIndex )
		maxLin = self.allMaxLins[ sheetIndex ]
		maxRow = self.allMaxCols[ sheetIndex ]

		colIni = 0 if ( not whatCels[0][1] or whatCels[0][1] <= 0 ) else whatCels[0][1]-1
		colEnd = maxRow if ( not whatCels[1][1] or whatCels[1][1] > maxRow ) else whatCels[1][1]
		linIni = 0 if ( not whatCels[0][0] or whatCels[0][0] <= 0 ) else whatCels[0][0]-1
		linEnd = maxLin if ( not whatCels[1][0] or whatCels[1][0] > maxLin ) else whatCels[1][0]

		selectedRows = [ r[ colIni : colEnd ] for r in allData ]
		selectedLines = selectedRows[ linIni : linEnd ]

		for sl in selectedLines:
			temp = []
			for s in sl:
				if s['tipo'] == 3: s['valor'] = self.parseExcelDate( s['valor'] )
				temp.append( s['valor'] )
			selectedData.append(temp)

		return selectedData
