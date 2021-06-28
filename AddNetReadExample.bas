Attribute VB_Name = "AddNetReadExample"
Option Explicit


Sub AddNetReadExample()

    'declare object Visum and other objects
    Dim Visum As Object
    Dim AddNetReadController As Object
    Dim NetReadRouteSearchTSysController As Object
    Dim NetReadRouteSearchController As Object
    
    'create the Visum-Object and load version, filename from cell B3
    Set Visum = CreateObject("Visum.Visum.180")
    Visum.LoadVersion Cells(5, 2)

    'create AddNetRead-Object and specify desired conflict avoiding method
    Set AddNetReadController = Visum.CreateAddNetReadController
    AddNetReadController.SetConflictAvoidingForAll 10000, "tra_"

    'create NetRouteSearchTSys-Object and choose route search options
    'create one object per TSys if desired
    Set NetReadRouteSearchTSysController = Visum.CreateNetReadRouteSearchTSys
    NetReadRouteSearchTSysController.DontRead

    'create NetRouteSearch-Object and assign NetRouteSearchTSys-objects
    Set NetReadRouteSearchController = Visum.CreateNetReadRouteSearch
    NetReadRouteSearchController.SetForAllTSys NetReadRouteSearchTSysController

    'additively read the net file, filename from cell B4
    Visum.LoadNet Cells(6, 2), True, NetReadRouteSearchController, AddNetReadController

    'write version file, filename from cell B5
    Visum.SaveVersion Cells(7, 2)
  
    'delete all objects to close VISUM-software
    Set Visum = Nothing
    
End Sub


