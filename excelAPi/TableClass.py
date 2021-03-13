from excelAPI import *

class DefinedSpace:
   def __init__(self,name,sheet,coordinate):
      self.coord =  coordinate
      self.name = name
      self.sheet = sheet

   def Get_Axis(self):
      return Get_numberic_axis(self.coord)

   def Update(self):
      pass

   def __str__(self) -> str:
      return f"{self.name} on {self.sheet} at {self.coord}"


class TitleSpace(DefinedSpace):
   def __init__(self,name,sheet,coordinate):
      super().__init__(name, sheet, coordinate)
      self.string = self.sheet[self.coord].value
   
   def Get_title(self):
      return self.string

   def Set_title(self,string):
      self.string = string
      self.sheet[self.coord] = string

   def __str__(self):
      return "TitleSpace " + super().__str__()

         
class DescriptionSpace(DefinedSpace):
   def __init__(self,name,sheet,coordinate):
      super().__init__(name, sheet, coordinate)
      start,end = Get_numberic_axis(coordinate)
       
      self.key_name = Get_cell_value(self.sheet,column=start[1],row=start[0])
      self.key_value = Get_cell_value(self.sheet,column=start[1]+2,row=start[0])
   
   def Set_key_name(self,new_key_name):
      start,end = Get_numberic_axis(self.coord)
      self.key_name = new_key_name
      Set_cell_value(self.sheet,new_key_name,column=start[1],row=start[0])

   def Set_key_value(self,new_key_value):
      start,end = Get_numberic_axis(self.coord)
      self.key_value = new_key_value
      Set_cell_value(self.sheet,new_key_value,column=start[1]+2,row=start[0])

   def Get_key_name(self):
      return self.key_name

   def Get_key_value(self):
      return self.key_value

   def __str__(self):   
      return "DescSpace " + super().__str__()


class TableSpace(DefinedSpace):
   
   def __init__(self,name,sheet,coordinate):
      super().__init__(name, sheet, coordinate)
      start,end = Get_numberic_axis(coordinate)
      self.keys_list= list()
      for index in range(start[1],end[1]+1):
         key = Get_cell_value(self.sheet, column=index, row=start[0])
         self.keys_list.append(key)
      row = start[0]+1
      while Get_cell_value(self.sheet, column=start[1], row=row) is not None:
         row += 1
      self.empty_space = [row,start[1]]
      
   def Get_dataTable(self):
      return_list = list()
      start,end = Get_numberic_axis(self.coord)
      for row in range(start[0]+1,self.empty_space[0]):
         temporary_dict = dict()
         i = start[1]
         for key in self.keys_list:
            temporary_dict[key] = Get_cell_value(self.sheet,column=i,row=row)
            i= i+1
         return_list.append(temporary_dict)
      return return_list
   
   def Add_new_row(self,data):
      row,start_column = self.empty_space
      end_column = start_column + len(self.keys_list)
      i = 0;
      for column in range(start_column,end_column):
         value = data[self.keys_list[i]]
         Set_cell_value(self.sheet,value,column=column,row=row)
         i += 1
      self.empty_space[0] += 1
      

   def Get_value_from_row(self,relative_row):
      start,end = Get_numberic_axis(self.coord)
      if relative_row < self.empty_space[0]:
         return_dict = dict()
         row,start_column = self.empty_space
         end_column = start_column + len(self.keys_list)
         i = 0;
         for column in range(start_column,end_column):
            value = Get_cell_value(self.sheet,column=column,row=start[0]+relative_row)
            key = self.keys_list[i]
            return_dict[key] = value
            i += 1
         return return_dict

   def Change_row_value(self,relative_row,data):
      start,end = Get_numberic_axis(self.coord) 
      if relative_row < self.empty_space[0]:
         row,column = self.empty_space
         for key in data.keys():
            value = data[key]
            index = self.keys_list.index(key)
            Set_cell_value(self.sheet,value,column=column+index,row=start[0]+relative_row)

   def __str__(self):
      return "TableSpace " + super().__str__()       


def Create_table_from_definedSpace(workbook):
   Title_list = list()
   Table_list = list()
   Desc_list = list()
   Space_dict = Load_definedname(workbook)
   if len(Space_dict) > 0:
      Title_pattern = re.compile(r"^TitleSpace[.].+")
      Table_pattern = re.compile(r"^TableSpace[.].+")
      Desc_pattern = re.compile(r"^DescSpace[.].+")
      for key in Space_dict.keys():
         name = key
         sheet = Space_dict[key][0]
         coord = Space_dict[key][1]
         if Title_pattern.search(key):
            table = TitleSpace(key,sheet,coord)
            Title_list.append(table)
         elif Table_pattern.search(key):
            table = TableSpace(key,sheet,coord)
            Table_list.append(table)      
         elif Desc_pattern.search(key):
            table = DescriptionSpace(key,sheet,coord)
            Desc_list.append(table)
   return Title_list,Desc_list,Table_list

        




   

