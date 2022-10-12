import os
import win32com.client as win32
#import PIconnect as PI
import numpy as np
#import matplotlib.pyplot as plt
#import time
import pandas as pd
#from IPython.display import clear_output
#import plotly.graph_objects as go


class HyCase:
    def __init__(self, path):
        '''
        Connection to the HYSYS simulation case
        Defining the most important COMobjects on the flowsheet
        '''
        
        self.app = win32.Dispatch("HYSYS.Application")
        if path == "Active":
            self.case = self.app.ActiveDocument
        else:
            file = os.path.abspath(path)
            self.case = self.app.SimulationCases.Open(file)
            
        self.file_name      = self.case.Title.Value
        self.thermo_package = self.case.Flowsheet.FluidPackage.PropertyPackageName
        self.comp_list      = np.array([i.name for i in self.case.Flowsheet.FluidPackage.Components])
        self.MatStreamList  = [i.name for i in self.case.Flowsheet.MaterialStreams]
        self.MatStreams     = self.case.Flowsheet.MaterialStreams
        self.EnerStreamList = [i.name for i in self.case.Flowsheet.EnergyStreams]
        self.OpList         = [i.name for i in self.case.Flowsheet.Operations]
        self.Solver         = self.case.Solver
        
    def set_visible(self, visibility = 0):
        """
        Determines  whether the HYSYS window visible or not

        Args:
            visibility (int) = 0 (if invisible, default value)
            visibility (int) = 1 (if visible)
        """        
        self.case.Visible = visibility
    
    def get_stream_compositions(self):
        components = self.case.Flowsheet.FluidPackage.Components.Names
        lenght = len(components)
        names = np.array(components).reshape(lenght,1)
        stream_names = np.array(['Component names'])
        
        for stream in self.MatStreams.Names:
            ms = self.MatStreams.Item(stream)
            comp = ms.ComponentMassFractionValue
            comp = np.around(comp,4)
            frac = pd.DataFrame(comp)
            names = np.append(names,frac, axis =1)
            stream_names = np.append(stream_names, stream)
        
        np.set_printoptions(precision=4)
        stream_data = pd.DataFrame(names, columns = stream_names)
        
        print('Process streams in the simulation case:')
        print(stream_names)
        print('Number of material streams:', len(stream_names)-1)
        return stream_data
    
    def close(self):
        self.case.Close()
        self.app.quit()
        del self
    
    def save(self):
        self.case.Save()
    
    def __str__(self) -> str:
        """
        Prints the basic information about the current flowsheet.
        """        
        return f"File: {self.file_name}\n Thermodynamical package: {self.thermo_package}\n Number of components: {len(self.comp_list)}\n Number of material streams: {len(self.MatStreams)}"


class ProcessStream:
    '''
    Superclass for all process stream in the HYSYS simulation
    '''
    
    def __init__ (self, COMObject):
        self.COMObject   = COMObject
        self.connections = self.get_connections()
        self.name        = self.COMObject.name
        
    def get_connections(self) -> dict:
        """Stores the connections of the process stream into a dictionary.

        Returns:
            dict: Returns a dictionary of the upstream and downstream connections
        """        
        upstream   = [i.name for i in self.COMObject.UpstreamOpers]
        downstream = [i.name for i in self.COMObject.DownstreamOpers]
        return {"Upstream": upstream, "Downstream": downstream}


class MaterialStream(ProcessStream):
    def __init__(self, COMObject, comp_list):
        """Reads the COMObjectfrom the simulation. This class designs a material stream, which has a series
        of properties.

        Args:
            COMObject (COMObject): HYSYS COMObject.
            comp_list (list): List of components in the process stream.
        """        
        super().__init__(COMObject)
        self.comp_list   = comp_list
        self.MassFlow    = COMObject.MassFlow.GetValue('kg/h')
        self.Temperature = COMObject.Temperature.GetValue('C')
        self.Pressure    = COMObject.Pressure.GetValue('bar')
        self.composition = pd.DataFrame(np.vstack(COMObject.ComponentMassFractionValue), index = comp_list, columns = ['Wt fraction'])
        


case = HyCase(r"C:\Users\arosomogyi\OneDrive - MOLGROUP\Docu\Process Simulation\2022\Aromás\Aromatic Unit  Model\101 col analysis\refctification_sec.hsc")
case.set_visible(1)