Attribute VB_Name = "modLE"
Option Explicit
'caractéristique d'une section
Private Enum SectionCharacteristics
IMAGE_SCN_TYPE_REG = &H0
IMAGE_SCN_TYPE_DSECT = &H1
IMAGE_SCN_TYPE_NOLOAD = &H2
IMAGE_SCN_TYPE_GROUP = &H4
IMAGE_SCN_TYPE_NO_PAD = &H8
IMAGE_SCN_TYPE_COPY = &H10
IMAGE_SCN_CNT_CODE = &H20
IMAGE_SCN_CNT_INITIALIZED_DATA = &H40
IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80
IMAGE_SCN_LNK_OTHER = &H100
IMAGE_SCN_LNK_INFO = &H200
IMAGE_SCN_TYPE_OVER = &H400
IMAGE_SCN_LNK_REMOVE = &H800
IMAGE_SCN_LNK_COMDAT = &H1000
IMAGE_SCN_MEM_FARDATA = &H8000
IMAGE_SCN_MEM_PURGEABLE = &H20000
IMAGE_SCN_MEM_16BIT = &H20000
IMAGE_SCN_MEM_LOCKED = &H40000
IMAGE_SCN_MEM_PRELOAD = &H80000
IMAGE_SCN_ALIGN_1BYTES = &H100000
IMAGE_SCN_ALIGN_2BYTES = &H200000
IMAGE_SCN_ALIGN_4BYTES = &H300000
IMAGE_SCN_ALIGN_8BYTES = &H400000
IMAGE_SCN_ALIGN_16BYTES = &H500000
IMAGE_SCN_ALIGN_32BYTES = &H600000
IMAGE_SCN_ALIGN_64BYTES = &H700000
IMAGE_SCN_ALIGN_128BYTES = &H800000
IMAGE_SCN_ALIGN_256BYTES = &H900000
IMAGE_SCN_ALIGN_512BYTES = &HA00000
IMAGE_SCN_ALIGN_1024BYTES = &HB00000
IMAGE_SCN_ALIGN_2048BYTES = &HC00000
IMAGE_SCN_ALIGN_4096BYTES = &HD00000
IMAGE_SCN_ALIGN_8192BYTES = &HE00000
IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000
IMAGE_SCN_MEM_DISCARDABLE = &H2000000
IMAGE_SCN_MEM_NOT_CACHED = &H4000000
IMAGE_SCN_MEM_NOT_PAGED = &H8000000
IMAGE_SCN_MEM_SHARED = &H10000000
IMAGE_SCN_MEM_EXECUTE = &H20000000
IMAGE_SCN_MEM_READ = &H40000000
IMAGE_SCN_MEM_WRITE = &H80000000
End Enum

'*****************************************************************************
'*                                                                           *
'*    This structure define format of LE header for OS/2,Windows exe files   *
'*       ----------------------------------------------------------          *
'*                                                                           *
'*    Author Trigub Serge. B&M&T Corp.                                       *
'*           10 January 1993                                                 *
'*                                                                           *
'*****************************************************************************
'
Private Type LE_Header_define
LE_Signature                            As Integer    ' Signature 'LE' for exe header
LE_Byte_Order                           As Byte    '
LE_Word_Order                           As Byte    '
LE_Exec_Format_Level                    As Long    '
LE_CPU_Type                             As Integer    '
LE_Target_OS                            As Integer    '
LE_Module_Version                       As Long    '
LE_Module_Type_Flags                    As Long    '
LE_Number_Of_Memory_Pages               As Long    '
LE_Initial_CS                           As Long    '
LE_Initial_EIP                          As Long    '
LE_Initial_SS                           As Long    '
LE_Initial_ESP                          As Long    '
LE_Memory_Page_Size                     As Long    '
LE_Bytes_On_Last_Page                   As Long    '
LE_Fixup_Section_Size                   As Long    '
LE_Fixup_Section_Checksum               As Long    '
LE_Loader_Section_Size                  As Long    '
LE_Loader_Section_CheckSum              As Long    '
LE_Object_Table_Offset                  As Long    '
LE_Object_Table_Entries                 As Long    '
LE_Object_Page_Map_Table_Offset         As Long    '
LE_Object_Iterate_Data_Map_Offset       As Long    '
LE_Resource_Table_Offset                As Long    '
LE_Resource_Table_Entries               As Long    '
LE_Resident_Names_Table_Offset          As Long    '
LE_Entry_Table_Offset                   As Long    '
LE_Module_Directives_Table_Offset       As Long    '
LE_Module_Directives_Table_Entries      As Long    '
LE_Fixup_Page_Table_Offset              As Long    '
LE_Fixup_Record_Table_Offset            As Long    '
LE_Imported_Module_Names_Table_Offset   As Long    '
LE_Imported_Modules_Count               As Long    '
LE_Imported_Procedure_Name_Table_Offset As Long    '
LE_Per_page_Checksum_Table_Offset       As Long    '
LE_First_Pages_Offset                    As Long    '
LE_Preload_Page_Count                   As Long    '
LE_Nonresident_Names_Table_Offset       As Long    '
LE_Nonresident_Names_Table_Length       As Long    '
LE_Nonresident_Names_Table_Checksum     As Long    '
LE_Automatic_Data_Object                As Long    '
LE_Debug_Information_Offset             As Long    '
LE_Debug_Information_Length             As Long    '
LE_Preload_Instance_Pages_Number        As Long    '
LE_Demand_Instance_Pages_Number         As Long    '
LE_Extra_Heap_Allocation                As Long    '
LE_Unknown                              As Long    '
'
End Type
'_______________________________________________________________________________
'
'LE_Module_Type_Flags_Define             RECORD  {
'
'                          LE_EXE_Module_Is_DLL          :1
'                          LE_EXE_Reserved1              :1
'                          LE_EXE_Errors_In_Module       :1
'                          LE_EXE_Reserved2              :1
'                          LE_EXE_Code_Load_Application  :1
'                          LE_EXE_Application_Type       :3
'                          LE_EXE_Reserved3              :2
'                          LE_EXE_No_External_FIXUP      :1
'                          LE_EXE_No_Internal_FIXUP      :1
'                          LE_EXE_Protected_Mode_Only    :1
'                          LE_EXE_Global_Initialization  :1
'                          LE_EXE_Multipledata           :1
'                          LE_EXE_Singledata             :1
'
'                                                }
'
'-------------------------------------------------------------------------------
'
Private Type LE_Object_Table_Define
'
LE_OBJ_Virtual_Segment_Size             As Long
LE_OBJ_Relocation_Base_Address          As Long
LE_OBJ_FLAGS                            As Long
LE_OBJ_Page_MAP_Index                   As Long
LE_OBJ_Page_MAP_Entries                 As Long
LE_OBJ_Name(3)                          As Byte
'
End Type
'_______________________________________________________________________________
'
Private Enum LE_OBJ_FLAGS_Define
                        LE_OBJ_FL_I_O_Privilage_Level = &H8000&
                        LE_OBJ_FL_Conforming_Segment = &H4000&
                        LE_OBJ_FL_BIG_Segment = &H2000&
                        LE_OBJ_FL_16_16_Alias = &H1000&
                        LE_OBJ_FL_Reserved = &H800&
                        LE_OBJ_FL_Resident_Long_Locable = &H400&
                        LE_OBJ_FL_Segment_Type_ResidentContigous = &H300&
                        LE_OBJ_FL_Segment_Type_Resident = &H200&
                        LE_OBJ_FL_Segment_Type_ZeroFilled = &H100&
                        LE_OBJ_FL_Segment_Type_Normal = &H0&
                        LE_OBJ_FL_Segment_Invalid = &H80&
                        LE_OBJ_FL_Segment_Preloaded = &H40&
                        LE_OBJ_FL_Segment_Shared = &H20&
                        LE_OBJ_FL_Segment_Discardable = &H10&
                        LE_OBJ_FL_Segment_Resource = &H8&
                        LE_OBJ_FL_Segment_Executable = &H4&
                        LE_OBJ_FL_Segment_Writable = &H2&
                        LE_OBJ_FL_Segment_Readable = &H1&
End Enum
Private Enum LE_OBJ_FL_Segment_Type_ENUM
'
                LE_OBJ_FL_Segment_Type_Normal = 0
                LE_OBJ_FL_Segment_Zero_Filled
                LE_OBJ_FL_Segment_Resident
                LE_OBJ_FL_Segment_Resident_contiguous
'
End Enum
'
'-------------------------------------------------------------------------------
'
Private Type LE_Page_Map_Table_Define
'
LE_PM_High_Page_Number          As Integer
LE_PM_Low_Page_Number           As Byte
LE_PM_FLAGS                     As Byte
'
End Type
'
'-------------------------------------------------------------------------------
'
'LE_PM_FLAGS_Define      RECORD  {
'                        LE_PM_FLG_Page_Type     :2
'                        LE_PM_FLG_Reserved      :6
'                        LE_PM_FLG_End_Page      :2
'
'                                }
'_______________________________________________________________________________
'
Private Enum LE_PM_FLG_Page_Type_Enum
                        LE_Legal_Page = 0
                        LE_Iterated_Page = 1
                        LE_Invalid_Page = 2
                        LE_Zero_Filled_Page = 3
End Enum
'_______________________________________________________________________________
'
Private Type LE_Entry_Define
'
LE_Entry_Entry_Flags            As Byte
'
'                                Union
'LE_Entry_Word_Offset            As Integer
LE_Entry_Dword_Offset           As Long
'                                ENDS
'
End Type
'
Private Type LE_Entry_Table_Define
'
LE_Entry_Number_of_Entries      As Byte
LE_Entry_Bungle_Flags           As Byte
LE_Entry_Object_Index           As Integer
LE_Entry_First_Entry() As LE_Entry_Define
'
End Type
'_______________________________________________________________________________
'
'
'-------------------------------------------------------------------------------
'
Private Enum LE_Entry_Bungle_Flags_define
                        LE_EB_32_Bits_Entry = &H2&
                        LE_EB_32_Valid_Entry = &H1&
End Enum
'_______________________________________________________________________________
'
Private Type LE_Fixup_Record_Table_Define
'
LE_Fixup_Relocation_Address_Type                        As Byte
LE_Fixup_Relocation_Type                                As Byte
LE_Fixup_Relocation_Page_Offset                        As Long
LE_Fixup_Segment_or_Module_Index                        As Byte
LE_Fixup_Offset_Or_Ordinal_Value                        As Long
LE_Fixup_AddValue_                                      As Long
LE_Fixup_Extra_                                         As Integer
LE_Fixup_Offset_                                        As Long
LE_Fixup_Offset_Counter                                 As Byte
'
End Type
'_______________________________________________________________________________
'
Private Enum LE_Rel_Addr_Type_Define
                        LE_RAT_List_Offset = &H20
                        LE_RAT_Fixup16_16_Alias = &H10
                        LE_RAT_Rel_Addr_Type = &HF&
End Enum
'_______________________________________________________________________________
'
Private Enum LE_Relocation_Address_Type_ENUM
                        LE_RA_Low_Byte = 0
                        LE_RA_16_bits_selector = 2
                        LE_RA_32_bits_Far_Pointer = 3
                        LE_RA_16_bits_Offset = 5
                        LE_RA_48_bits_Far_Pointer = 6
                        LE_RA_32_bits_Offset = 7
                        LE_RA_32_bits_EIP_Rel = 8
End Enum
'_______________________________________________________________________________
'
Private Enum LE_Reloc_Type_Define
                        LE_RT_Ordinal_Byte = &H80&
                        LE_RT_Reserv1 = &H40&
                        LE_RT_ABS_Dword = &H20&
                        LE_RT_Target_Offset_32 = &H10&
                        LE_RT_Reserv2 = &H8&
                        LE_RT_ADDITIVE_Type = &H4&
                        LE_RT_Reloc_Type = &H3&
End Enum
'_______________________________________________________________________________
'
Private Enum LE_Relocation_Type_ENUM
                        LE_RT_Internal_Reference = 0
                        LE_RT_Imported_Ordinal = 1
                        LE_RT_Imported_Name = 2
                        LE_RT_OS_FIXUP = 3
End Enum
'_______________________________________________________________________________
'
Private Enum LE_CPU_Type_ENUM
'
                        LE_CPU_i80286 = &H1
                        LE_CPU_i80386 = &H2
                        LE_CPU_i80486 = &H3
                        LE_CPU_i80586 = &H4
                        LE_CPU_i860_N10 = &H20
                        LE_CPU_i860_N11 = &H21
                        LE_CPU_MIPS_Mark_I = &H40
                        LE_CPU_MIPS_Mark_II = &H41
                        LE_CPU_MIPS_Mark_III = &H42
'
End Enum

Private Type VxD_Desc_Block
        DDB_Next                As Long            ' VMM RESERVED FIELD
        DDB_SDK_Version         As Integer         ' VMM RESERVED FIELD
        DDB_Req_Device_Number   As Integer         ' Required device number
        DDB_Dev_Major_Version   As Byte            ' Major device number
        DDB_Dev_Minor_Version   As Byte            ' Minor device number
        DDB_Flags               As Integer         ' Flags for init calls complete
        DDB_Name(7) As Byte                        ' Device name
        DDB_Init_Order          As Long            ' Initialization Order
        DDB_Control_Proc        As Long            ' Offset of control procedure
        DDB_V86_API_Proc        As Long            ' Offset of API procedure (or 0)
        DDB_PM_API_Proc         As Long            ' Offset of API procedure (or 0)
        DDB_V86_API_CSIP        As Long            ' CS:IP of API entry point
        DDB_PM_API_CSIP         As Long            ' CS:IP of API entry point
        DDB_Reference_Data      As Long            ' Reference data from real mode
        DDB_Service_Table_Ptr   As Long            ' Pointer to service table
        DDB_Service_Table_Size  As Long            ' Number of services
        DDB_Win32_Service_Table As Long            ' Pointer to Win32 services
        DDB_Prev_0              As Long            ' Pointer to previous DDB
        DDB_Size_0              As Long            ' Size of VxD_Desc_Block
End Type
Dim arrEntryTable() As LE_Entry_Table_Define
Dim EntryPointsCol As Collection
Dim InitEntryPoint As Long
Private strModuleDescription As String, strVXDDesc As String, strModuleName As String
Private strVMMCalls()
Private strVMMCall0001(), strVMMCall0003(), strVMMCall0004(), strVMMCall0005(), strVMMCall0006(), strVMMCall0007(), strVMMCall000C(), strVMMCall000D(), strVMMCall000E(), strVMMCall0010(), strVMMCall0011(), strVMMCall0012(), strVMMCall0014(), strVMMCall0015(), strVMMCall0017(), strVMMCall0018(), strVMMCall001A(), strVMMCall001B(), strVMMCall001C(), strVMMCall0020(), strVMMCall0021(), strVMMCall0026(), strVMMCall0027(), strVMMCall0028(), strVMMCall002A(), strVMMCall002B(), strVMMCall0033(), strVMMCall0036(), strVMMCall0037(), strVMMCall0038(), strVMMCall0040(), strVMMCall0043(), strVMMCall0048(), strVMMCall004A(), strVMMCall004B()

Private Sub InitVMMCalls()
strVMMCall0001 = Array( _
                    "Get_VMM_Version", "Get_Cur_VM_Handle", "Test_Cur_VM_Handle", "Get_Sys_VM_Handle", "Test_Sys_VM_Handle", "Validate_VM_Handle", "Get_VMM_Reenter_Count", "Begin_Reentrant_Execution", "End_Reentrant_Execution", "Install_V86_Break_Point", "Remove_V86_Break_Point", "Allocate_V86_Call_Back", "Allocate_PM_Call_Back", "Call_When_VM_Returns", "Schedule_Global_Event", "Schedule_VM_Event", "Call_Global_Event", "Call_VM_Event", "Cancel_Global_Event", "Cancel_VM_Event", "Call_Priority_VM_Event", "Cancel_Priority_VM_Event", "Get_NMI_Handler_Addr", "Set_NMI_Handler_Addr", "Hook_NMI_Event", "Call_When_VM_Ints_Enabled", "Enable_VM_Ints", "Disable_VM_Ints", "Map_Flat", "Map_Lin_To_VM_Addr", "Adjust_Exec_Priority", _
                    "Begin_Critical_Section", "End_Critical_Section", "End_Crit_And_Suspend", "Claim_Critical_Section", "Release_Critical_Section", "Call_When_Not_Critical", "Create_Semaphore", "Destroy_Semaphore", "Wait_Semaphore", "Signal_Semaphore", "Get_Crit_Section_Status", "Call_When_Task_Switched", "Suspend_VM", "Resume_VM", "No_Fail_Resume_VM", "Nuke_VM", "Crash_Cur_VM", "Get_Execution_Focus", "Set_Execution_Focus", "Get_Time_Slice_Priority", "Set_Time_Slice_Priority", "Get_Time_Slice_Granularity", "Set_Time_Slice_Granularity", "Get_Time_Slice_Info", "Adjust_Execution_Time", "Release_Time_Slice", "Wake_Up_VM", "Call_When_Idle", "Get_Next_VM_Handle", "Set_Global_Time_Out", _
                    "Set_VM_Time_Out", "Cancel_Time_Out", "Get_System_Time", "Get_VM_Exec_Time", "Hook_V86_Int_Chain", "Get_V86_Int_Vector", "Set_V86_Int_Vector", "Get_PM_Int_Vector", "Set_PM_Int_Vector", "Simulate_Int", "Simulate_Iret", "Simulate_Far_Call", "Simulate_Far_Jmp", "Simulate_Far_Ret", "Simulate_Far_Ret_N", "Build_Int_Stack_Frame", "Simulate_Push", "Simulate_Pop", "_HeapAllocate", "_HeapReallocate", "_HeapFree", "_HeapGetSize", "_PageAllocate", "_PageReallocate", "_PageFree", "_PageLock", "_PageUnlock", "_PageGetSizeAddr", "_PageGetAllocInfo", "_GetFreePageCount", _
                    "_GetSysPageCount", "_GetVMpgCount", "_MapIntov86", "_PhysIntov86", "_TestGlobalv86mem", "_ModifyPageBits", "_CopyPageTable", "_LinMapIntov86", "_LinPageLock", "_LinPageUnlock", "_SetResetv86pageable", "_Getv86pageableArray", "_PageCheckLinRange", "_PageOutDirtyPages", "_PageDiscardPages", "_GetNulPageHandle", "_GetFirstv86page", "_MapPhysToLinear", "_GetAppFlatDSalias", "_SelectorMapFlat", "_GetDemandPageInfo", "_GetSetPageOutCount", "Hook_v86_Page", "_Assign_Device_v86_Pages", "_Deassign_Device_v86_Pages", "_Get_Device_v86_Pages_Array", "Mmgr_SetNulPageAddr", "_Allocate_GDT_Selector", "_Free_GDT_Selector", "_Allocate_LDT_Selector", _
                    "_Free_LDT_Selector", "_BuildDescriptorDwords", "_GetDescriptor", "_SetDescriptor", "_Mmgr_Toggle_HMA", "Get_Fault_Hook_Addrs", "Hook_v86_Fault", "Hook_PM_Fault", "Hook_VMM_Fault", "Begin_Nest_v86_Exec", "Begin_Nest_Exec", "Exec_Int", "Resume_Exec", "End_Nest_Exec", "Allocate_PM_App_CB_Area", "Get_Cur_PM_App_CB", "Set_v86_Exec_Mode", "Set_PM_Exec_Mode", "Begin_Use_Locked_PM_Stack", "End_Use_Locked_PM_Stack", "Save_Client_State", "Restore_Client_State", "Exec_VXD_Int", "Hook_Device_Service", "Hook_Device_v86_Api", "Hook_Device_PM_Api", "System_Control", "Simulate_IO", "Install_Mult_IO_Handlers", "Install_IO_Handler", _
                    "Enable_Global_Trapping", "Enable_Local_Trapping", "Disable_Global_Trapping", "Disable_Local_Trapping", "List_Create", "List_Destroy", "List_Allocate", "List_Attach", "List_Attach_Tail", "List_Insert", "List_Remove", "List_Deallocate", "List_Get_First", "List_Get_Next", "List_Remove_First", "_AddInstanceItem", "_Allocate_Device_CB_Area", "_Allocate_Global_v86_Data_Area", "_Allocate_Temp_v86_Data_Area", "_Free_Temp_v86_Data_Area", "Get_Profile_Decimal_Int", "Convert_Decimal_String", "Get_Profile_Fixed_Point", "Convert_Fixed_Point_String", "Get_Profile_Hex_Int", "Convert_Hex_String", "Get_Profile_Boolean", "Convert_Boolean_String", "Get_Profile_String", "Get_Next_Profile_String", _
                    "Get_Environment_String", "Get_Exec_Path", "Get_Config_Directory", "Openfile", "Get_PSP_Segment", "GetDOSvectors", "Get_Machine_Info", "GetSet_HMA_Info", "Set_System_Exit_Code", "Fatal_Error_Handler", "Fatal_Memory_Error", "Update_System_Clock", "Test_Debug_Installed", "Out_Debug_String", "Out_Debug_Chr", "In_Debug_Chr", "Debug_Convert_Hex_Binary", "Debug_Convert_Hex_Decimal", "Debug_Test_Valid_Handle", "Validate_Client_Ptr", "Test_Reenter", "Queue_Debug_String", "Log_Proc_Call", "Debug_Test_Cur_VM", "Get_PM_Int_Type", "Set_PM_Int_Type", "Get_Last_Updated_System_Time", "Get_Last_Updated_VM_Exec_Time", "Test_DBCS_Lead_Byte", "_AddFreePhysPage", _
                    "_PageResetHandlePAddr", "_SetLastV86Page", "_GetLastV86Page", "_MapFreePhysReg", "_UnmapFreePhysReg", "_XchgFreePhysReg", "_SetFreePhysRegCalBk", "Get_Next_Arena", "Get_Name_Of_Ugly_TSR", "Get_Debug_Options", "Set_Physical_HMA_Alias", "_GetGlblRng0V86IntBase", "_Add_Global_V86_Data_Area", "GetSetDetailedVMError", "Is_Debug_Chr", "Clear_Mono_Screen", "Out_Mono_Chr", "Out_Mono_String", "Set_Mono_Cur_Pos", "Get_Mono_Cur_Pos", "Get_Mono_Chr", "Locate_Byte_In_ROM", "Hook_Invalid_Page_Fault", "Unhook_Invalid_Page_Fault", "Set_Delete_On_Exit_File", "Close_VM", "Enable_Touch_1st_Meg", "Disable_Touch_1st_Meg", "Install_Exception_Handler", "Remove_Exception_Handler", _
                    "Get_Crit_Status_No_Block", "_GetLastUpdatedThreadExecTime", "_Trace_Out_Service", "_Debug_Out_Service", "_Debug_Flags_Service ; Assert conditions", "VMMAddImportModuleName", "VMM_Add_DDB", "VMM_Remove_DDB", "Test_VM_Ints_Enabled", "_BlockOnID", "Schedule_Thread_Event", "Cancel_Thread_Event", "Set_Thread_Time_Out", "Set_Async_Time_Out", "_AllocateThreadDataSlot", "_FreeThreadDataSlot", "_CreateMutex", "_DestroyMutex", "_GetMutexOwner", "Call_When_Thread_Switched", "VMMCreateThread", "_GetThreadExecTime", "VMMTerminateThread", "Get_Cur_Thread_Handle", "Test_Cur_Thread_Handle", "Get_Sys_Thread_Handle", "Test_Sys_Thread_Handle", "Validate_Thread_Handle", "Get_Initial_Thread_Handle", "Test_Initial_Thread_Handle", _
                    "Debug_Test_Valid_Thread_Handle", "Debug_Test_Cur_Thread", "VMM_GetSystemInitState", "Cancel_Call_When_Thread_Switched", "Get_Next_Thread_Handle", "Adjust_Thread_Exec_Priority", "_Deallocate_Device_CB_Area", "Remove_IO_Handler", "Remove_Mult_IO_Handlers", "Unhook_V86_Int_Chain", "Unhook_V86_Fault", "Unhook_PM_Fault", "Unhook_VMM_Fault", "Unhook_Device_Service", "_PageReserve", "_PageCommit", "_PageDecommit", "_PagerRegister", "_PagerQuery", "_PagerDeregister", "_ContextCreate", "_ContextDestroy", "_PageAttach", "_PageFlush", "_SignalID", "_PageCommitPhys", "_Register_Win32_Services", "Cancel_Call_When_Not_Critical", "Cancel_Call_When_Idle", "Cancel_Call_When_Task_Switched", _
                    "_Debug_Printf_Service", "_EnterMutex", "_LeaveMutex", "Simulate_VM_IO", "Signal_Semaphore_No_Switch", "_ContextSwitch", "_PageModifyPermissions", "_PageQuery", "_EnterMustComplete", "_LeaveMustComplete", "_ResumeExecMustComplete", "_GetThreadTerminationStatus", "_GetInstanceInfo", "_ExecIntMustComplete", "_ExecVxDIntMustComplete", "Begin_V86_Serialization", "Unhook_V86_Page", "VMM_GetVxDLocationList", "VMM_GetDDBList", "Unhook_NMI_Event", "Get_Instanced_V86_Int_Vector", "Get_Set_Real_DOS_PSP", "Call_Priority_Thread_Event", "Get_System_Time_Address", "Get_Crit_Status_Thread", "Get_DDB", "Directed_Sys_Control", "_RegOpenKey", "_RegCloseKey", "_RegCreateKey", _
                    "_RegDeleteKey", "_RegEnumKey", "_RegQueryValue", "_RegSetValue", "_RegDeleteValue", "_RegEnumValue", "_RegQueryValueEx", "_RegSetValueEx", "_CallRing3", "Exec_PM_Int", "_RegFlushKey", "_PageCommitContig", "_GetCurrentContext", "_LocalizeSprintf", "_LocalizeStackSprintf", "Call_Restricted_Event", "Cancel_Restricted_Event", "Register_PEF_Provider", "_GetPhysPageInfo", "_RegQueryInfoKey", "MemArb_Reserve_Pages", "Time_Slice_Sys_VM_Idle", "Time_Slice_Sleep", "Boost_With_Decay", "Set_Inversion_Pri", "Reset_Inversion_Pri", "Release_Inversion_Pri", "Get_Thread_Win32_Pri", "Set_Thread_Win32_Pri", "Set_Thread_Static_Boost", _
                    "Set_VM_Static_Boost", "Release_Inversion_Pri_ID", "Attach_Thread_To_Group", "Detach_Thread_From_Group", "Set_Group_Static_Boost", "_GetRegistryPath", "_GetRegistryKey", "Cleanup_Thread_State", "_RegRemapPreDefKey", "End_V86_Serialization", "_Assert_Range", "_Sprintf", "_PageChangePager", "_RegCreateDynKey", "_RegQueryMultipleValues", "Boost_Thread_With_VM", "Get_Boot_Flags", "Set_Boot_Flags", "_lstrcpyn", "_lstrlen", "_lmemcpy", "_GetVxDName", "Force_Mutexes_Free", "Restore_Forced_Mutexes", "_AddReclaimableItem", "_SetReclaimableItem", "_EnumReclaimableItem", "Time_Slice_Wake_Sys_VM", "VMM_Replace_Global_Environment", "Begin_Non_Serial_Nest_V86_Exec", _
                    "Get_Nest_Exec_Status", "Open_Boot_Log", "Write_Boot_Log", "Close_Boot_Log", "EnableDisable_Boot_Log", "_Call_On_My_Stack", "Get_Inst_V86_Int_Vec_Base", "_lstrcmpi", "_strupr", "Log_Fault_Call_Out", "_AtEventTime", "_PageOutPages", "_Call_On_My_Not_Flat_Stack", "_LinRegionLock", "_LinRegionUnlock", "_AttemptingSomethingDangerous", "_Vsprintf", "_Vsprintfw", "Load_FS_Service", "Assert_FS_Service", "ObsoleteRtlUnwind", "ObsoleteRtlRaiseException", "ObsoleteRtlRaiseStatus", "ObsoleteKeGetCurrentIrql", "ObsoleteKfRaiseIrql", "ObsoleteKfLowerIrql", "_Begin_Preemptable_Code", "_End_Preemptable_Code", "Set_Preemptable_Count", "ObsoleteKeInitializeDpc", _
                    "ObsoleteKeInsertQueueDpc", "ObsoleteKeRemoveQueueDpc", "HeapAllocateEx", "HeapReAllocateEx", "HeapGetSizeEx", "HeapFreeEx", "_Get_CPUID_Flags", "KeCheckDivideByZeroTrap", "_RegisterGARTHandler", "_GARTReserve", "_GARTCommit", "_GARTUnCommit", "_GARTFree", "_GARTMemAttributes", "KfRaiseIrqlToDpcLevel", "VMMCreateThreadEx", "_FlushCaches", "Set_Thread_Win32_Pri_BoYield", "_FlushMappedCacheBlock", "_ReleaseMappedCacheBlock", "Run_Preemptable_Events", "_MMPreSystemExit", "_MMPageFileShutDown", "_Set_Global_Time_Out_Ex", "Query_Thread_Priority")

strVMMCall0003 = Array( _
                    "VPICD_Get_Version", "VPICD_Virtualize_IRQ", "VPICD_Set_Int_Request", "VPICD_Clear_Int_Request", "VPICD_Phys_EOI", "VPICD_Get_Complete_Status", "VPICD_Get_Status", "VPICD_Test_Phys_Request", "VPICD_Physically_Mask", "VPICD_Physically_Unmask", "VPICD_Set_Auto_Masking", "VPICD_Get_IRQ_Complete_Status", "VPICD_Convert_Handle_To_IRQ", "VPICD_Convert_IRQ_To_Int", "VPICD_Convert_Int_To_IRQ", "VPICD_Call_When_Hw_Int", _
                    "VPICD_Force_Default_Owner", "VPICD_Force_Default_Behavior", "VPICD_Auto_Mask_At_Inst_Swap", "VPICD_Begin_Inst_Page_Swap", "VPICD_End_Inst_Page_Swap", "VPICD_Virtual_EOI", "VPICD_Get_Virtualization_Count", "VPICD_Post_Sys_Critical_Init", "VPICD_VM_SlavePIC_Mask_Change", "_VPICD_Clear_IR_Bits", "_VPICD_Get_Level_Mask", "_VPICD_Set_Level_Mask", "_VPICD_Set_Irql_Mask", "_VPICD_Set_Channel_Irql", "_VPICD_Prepare_For_Shutdown", _
                    "_VPICD_Register_Trigger_Handler")

strVMMCall0004 = Array( _
                    "VDMAD_Get_Version", "VDMAD_Virtualize_Channel", "VDMAD_Get_Region_Info", "VDMAD_Set_Region_Info", "VDMAD_Get_Virt_State", "VDMAD_Set_Virt_State", "VDMAD_Set_Phys_State", "VDMAD_Mask_Channel", "VDMAD_UnMask_Channel", "VDMAD_Lock_DMA_Region", "VDMAD_Unlock_DMA_Region", "VDMAD_Scatter_Lock", "VDMAD_Scatter_Unlock", "VDMAD_Reserve_Buffer_Space", "VDMAD_Request_Buffer", "VDMAD_Release_Buffer", _
                    "VDMAD_Copy_To_Buffer", "VDMAD_Copy_From_Buffer", "VDMAD_Default_Handler", "VDMAD_Disable_Translation", "VDMAD_Enable_Translation", "VDMAD_Get_EISA_Adr_Mode", "VDMAD_Set_EISA_Adr_Mode", "VDMAD_Unlock_DMA_Region_No_Dirty", "VDMAD_Phys_Mask_Channel", "VDMAD_Phys_Unmask_Channel", "VDMAD_Unvirtualize_Channel", "VDMAD_Set_IO_Address", "VDMAD_Get_Phys_Count", "VDMAD_Get_Phys_Status", "VDMAD_Get_Max_Phys_Page", _
                    "VDMAD_Set_Channel_Callbacks", "VDMAD_Get_Virt_Count", "VDMAD_Set_Virt_Count", "VDMAD_Get_Virt_Address", "VDMAD_Set_Virt_Address")

strVMMCall0005 = Array( _
                    "VTD_Get_Version", "VTD_Update_System_Clock", "VTD_Get_Interrupt_Period", "VTD_Begin_Min_Int_Period", "VTD_End_Min_Int_Period", "VTD_Disable_Trapping", "VTD_Enable_Trapping", "VTD_Get_Real_Time", "VTD_Get_Date_And_Time", "VTD_Adjust_VM_Count", "VTD_Delay", "VTD_GetTimeZoneBias", "ObsoleteKeQueryPerformanceCounter", "ObsoleteKeQuerySystemTime", "VTD_Install_IO_Handle", "VTD_Remove_IO_Handle", _
                    "_VTD_Delay_Ex", "VTD_Init_Timer")

strVMMCall0006 = Array( _
                    "V86MMGR_Get_Version", "V86MMGR_Allocate_V86_Pages", "V86MMGR_Set_EMS_XMS_Limits", "V86MMGR_Get_EMS_XMS_Limits", "V86MMGR_Set_Mapping_Info", "V86MMGR_Get_Mapping_Info", "V86MMGR_Xlat_API", "V86MMGR_Load_Client_Ptr", "V86MMGR_Allocate_Buffer", "V86MMGR_Free_Buffer", "V86MMGR_Get_Xlat_Buff_State", "V86MMGR_Set_Xlat_Buff_State", "V86MMGR_Get_VM_Flat_Sel", "V86MMGR_Map_Pages", "V86MMGR_Free_Page_Map_Region", "V86MMGR_LocalGlobalReg", _
                    "V86MMGR_GetPgStatus", "V86MMGR_SetLocalA20", "V86MMGR_ResetBasePages", "V86MMGR_SetAvailMapPgs", "V86MMGR_NoUMBInitCalls", "V86MMGR_Get_EMS_XMS_Avail", "V86MMGR_Toggle_HMA", "V86MMGR_Dev_Init", "V86MMGR_Alloc_UM_Page", "V86MMGR_Check_NHSupport")

strVMMCall0007 = Array( _
                    "PageSwap_Get_Version", "PageSwap_Test_Create", "PageSwap_Create", "PageSwap_Destroy", "PageSwap_In", "PageSwap_Out", "PageSwap_Test_IO_Valid", "PageSwap_Read_Or_Write", "PageSwap_Grow_File", "PageSwap_Init_File")

strVMMCall000C = Array( _
                    "VMD_Get_Version", "VMD_Set_Mouse_Type", "VMD_Get_Mouse_Owner", "VMD_Post_Pointer_Message", "VMD_Set_Cursor_Proc", "VMD_Call_Cursor_Proc", "VMD_Set_Mouse_Data", "VMD_Get_Mouse_Data", "VMD_Manipulate_Pointer_Message", "VMD_Set_Middle_Button", "VMD_Enable_Disable_Mouse_Events", "VMD_Post_Absolute_Pointer_Message")

strVMMCall000D = Array( _
                    "VKD_Get_Version", "VKD_Define_Hot_Key", "VKD_Remove_Hot_Key", "VKD_Local_Enable_Hot_Key", "VKD_Local_Disable_Hot_Key", "VKD_Reflect_Hot_Key", "VKD_Cancel_Hot_Key_State", "VKD_Force_Keys", "VKD_Get_Kbd_Owner", "VKD_Define_Paste_Mode", "VKD_Start_Paste", "VKD_Cancel_Paste", "VKD_Get_Msg_Key", "VKD_Peek_Msg_Key", "VKD_Flush_Msg_Key_Queue", "VKD_Enable_Keyboard", _
                    "VKD_Disable_Keyboard", "VKD_Get_Shift_State", "VKD_Filter_Keyboard_Input", "VKD_Put_Byte", "VKD_Set_Shift_State", "VKD_Send_Data", "VKD_Set_LEDs", "VKD_Set_Key_Rate", "VKD_Get_Key_Rate")

strVMMCall000E = Array( _
                    "VCD_Get_Version", "VCD_Set_Port_Global", "VCD_Get_Focus", "VCD_Virtualize_Port", "VCD_Acquire_Port", "VCD_Free_Port", "VCD_Acquire_Port_Windows_Style", "VCD_Free_Port_Windows_Style", "VCD_Steal_Port_Windows_Style", "VCD_Find_COM_Index", "VCD_Set_Port_Global_Special", "VCD_Virtualize_Port_Dynamic", "VCD_Unvirtualize_Port_Dynamic")

strVMMCall0010 = Array( _
                    "IOS_Get_Version", "IOS_BD_Register_Device", "IOS_Find_Int13_Drive", "IOS_Get_Device_List", "IOS_SendCommand", "IOS_BD_Command_Complete", "IOS_Synchronous_Command", "IOS_Register", "IOS_Requestor_Service", "IOS_Exclusive_Access", "IOS_Send_Next_Command", "IOS_Set_Async_Time_Out", "IOS_Signal_Semaphore_No_Switch", "IOSIdleStatus", "IOSMapIORSToI24", "IOSMapIORSToI21", _
                    "PrintLog", "IOS_deregister", "IOS_wait", "IOS_SpinDownDrives", "_IOS_query_udf_mount")

strVMMCall0011 = Array( _
                    "VMCPD_Get_Version", "VMCPD_Get_Virt_State", "VMCPD_Set_Virt_State", "VMCPD_Get_CR0_State", "VMCPD_Set_CR0_State", "VMCPD_Get_Thread_State", "VMCPD_Set_Thread_State", "_VMCPD_Get_FP_Instruction_Size", "VMCPD_Set_Thread_Precision", "VMCPD_Init_FP", "_KeSaveFloatingPointState", "_KeRestoreFloatingPointState", "VMCPD_Init_FP_State")

strVMMCall0012 = Array( _
                    "EBIOS_Get_Version", "EBIOS_Get_Unused_Mem")

strVMMCall0014 = Array( _
                    "VNETBIOS_Get_Version", "VNETBIOS_Register", "VNETBIOS_Submit", "VNETBIOS_Enum", "VNETBIOS_Deregister", "VNETBIOS_Register2", "VNETBIOS_Map", "VNETBIOS_Enum2")

strVMMCall0015 = Array( _
                    "DOSMGR_Get_Version", "_DOSMGR_Set_Exec_VM_Data", "DOSMGR_Copy_VM_Drive_State", "_DOSMGR_Exec_VM", "DOSMGR_Get_IndosPtr", "DOSMGR_Add_Device", "DOSMGR_Remove_Device", "DOSMGR_Instance_Device", "DOSMGR_Get_DOS_Crit_Status", "DOSMGR_Enable_Indos_Polling", "DOSMGR_BackFill_Allowed", "DOSMGR_LocalGlobalReg", "DOSMGR_Init_UMB_Area", "DOSMGR_Begin_V86_App", "DOSMGR_End_V86_App", "DOSMGR_Alloc_Local_Sys_VM_Mem", _
                    "DOSMGR_Grow_CDSs", "DOSMGR_Translate_Server_DOS_Call", "DOSMGR_MMGR_PSP_Change_Notifier")

strVMMCall0017 = Array( _
                    "SHELL_Get_Version", "SHELL_Resolve_Contention", "SHELL_Event", "SHELL_SYSMODAL_Message", "SHELL_Message", "SHELL_GetVMInfo", "_SHELL_PostMessage", "_SHELL_ShellExecute", "_SHELL_PostShellMessage", "SHELL_DispatchRing0AppyEvents", "SHELL_Hook_Properties", "SHELL_Unhook_Properties", "SHELL_Update_User_Activity", "_SHELL_QueryAppyTimeAvailable", "_SHELL_CallAtAppyTime", "_SHELL_CancelAppyTimeEvent", _
                    "_SHELL_BroadcastSystemMessage", "_SHELL_HookSystemBroadcast", "_SHELL_UnhookSystemBroadcast", "_SHELL_LocalAllocEx", "_SHELL_LocalFree", "_SHELL_LoadLibrary", "_SHELL_FreeLibrary", "_SHELL_GetProcAddress", "_SHELL_CallDll", "_SHELL_SuggestSingleMSDOSMode", "SHELL_CheckHotkeyAllowed", "_SHELL_GetDOSAppInfo", "_SHELL_Update_User_Activity_Ex")

strVMMCall0018 = Array( _
                    "VMPoll_Get_Version", "VMPoll_Enable_Disable", "VMPoll_Reset_Detection", "VMPoll_Check_Idle")

strVMMCall001A = Array( _
                    "DOSNET_Get_Version", "DOSNET_Send_FILESYSCHANGE", "DOSNET_Do_PSP_Adjust")

strVMMCall001B = Array( _
                    "VFD_Get_Version")

strVMMCall001C = Array( _
                    "VDD2_Get_Version")

strVMMCall0020 = Array( _
                    "Int13_Get_Version", "Int13_Device_Registered", "Int13_Translate_VM_Int", "Int13_Hooking_BIOS_Int", "Int13_Unhooking_BIOS_Int")

strVMMCall0021 = Array( _
                    "PageFile_Get_Version", "PageFile_Init_File", "PageFile_Clean_Up", "PageFile_Grow_File", "PageFile_Read_Or_Write", "PageFile_Cancel", "PageFile_Test_IO_Valid", "PageFile_Get_Size_Info", "PageFile_Set_Async_Manager", "PageFile_Call_Async_Manager")

strVMMCall0026 = Array( _
                    "_VPOWERD_Get_Version", "_VPOWERD_Get_APM_BIOS_Version", "_VPOWERD_Get_Power_Management_Level", "_VPOWERD_Set_Power_Management_Level", "_VPOWERD_Set_Device_Power_State", "_VPOWERD_Set_System_Power_State", "_VPOWERD_Restore_Power_On_Defaults", "_VPOWERD_Get_Power_Status", "_VPOWERD_Get_Power_State", "_VPOWERD_OEM_APM_Function", "_VPOWERD_Register_Power_Handler", "_VPOWERD_Deregister_Power_Handler", "_VPOWERD_W32_Get_System_Power_Status", "_VPOWERD_W32_Set_System_Power_State", "_VPOWERD_Get_Capabilities", "_VPOWERD_Enable_Resume_On_Ring", _
                    "_VPOWERD_Disable_Resume_On_Ring", "_VPOWERD_Set_Resume_Timer", "_VPOWERD_Get_Resume_Timer", "_VPOWERD_Disable_Resume_Timer", "_VPOWERD_Enable_Timer_Based_Requests", "_VPOWERD_Disable_Timer_Based_Requests", "_VPOWERD_W32_Get_Power_Status", "_VPOWERD_Get_Timer_Based_Requests_Status", "_VPOWERD_Get_Ring_Resume_Status", "_VPOWERD_Transfer_Control", "_VPOWERD_OS_Shutdown", "_VPOWERD_Indicate_User_Arrival", "_VPOWERD_Get_Battery_Unit_Status", "_VPOWERD_Get_Battery_Unit_Presence", "_VPOWERD_Disable_APM_Idle", _
                    "_VPOWERD_Get_Mode", "_VPOWERD_Critical_Shutdown")

strVMMCall0027 = Array( _
                    "VXDLDR_GetVersion", "VXDLDR_LoadDevice", "VXDLDR_UnloadDevice", "VXDLDR_DevInitSucceeded", "VXDLDR_DevInitFailed", "VXDLDR_GetDeviceList", "VXDLDR_UnloadMe", "_PELDR_LoadModule", "_PELDR_GetModuleHandle", "_PELDR_GetModuleUsage", "_PELDR_GetEntryPoint", "_PELDR_GetProcAddress", "_PELDR_AddExportTable", "_PELDR_RemoveExportTable", "_PELDR_FreeModule", "VXDLDR_Notify", _
                    "_PELDR_InitCompleted", "_PELDR_LoadModuleEx", "_PELDR_LoadModule2")

strVMMCall0028 = Array( _
                    "NdisGetVersion", "NdisAllocateSpinLock", "NdisFreeSpinLock", "NdisAcquireSpinLock", "NdisReleaseSpinLock", "NdisOpenConfiguration", "NdisReadConfiguration", "NdisCloseConfiguration", "NdisReadEisaSlotInformation", "NdisReadMcaPosInformation", "NdisAllocateMemory", "NdisFreeMemory", "NdisSetTimer", "NdisCancelTimer", "NdisStallExecution", "NdisInitializeInterrupt", _
                    "NdisRemoveInterrupt", "NdisSynchronizeWithInterrupt", "NdisOpenFile", "NdisMapFile", "NdisUnmapFile", "NdisCloseFile", "NdisAllocatePacketPool", "NdisFreePacketPool", "NdisAllocatePacket", "NdisReinitializePacket", "NdisFreePacket", "NdisQueryPacket", "NdisAllocateBufferPool", "NdisFreeBufferPool", "NdisAllocateBuffer", _
                    "NdisCopyBuffer", "NdisFreeBuffer", "NdisQueryBuffer", "NdisGetBufferPhysicalAddress", "NdisChainBufferAtFront", "NdisChainBufferAtBack", "NdisUnchainBufferAtFront", "NdisUnchainBufferAtBack", "NdisGetNextBuffer", "NdisCopyFromPacketToPacket", "NdisRegisterProtocol", "NdisDeregisterProtocol", "NdisOpenAdapter", "NdisCloseAdapter", "NdisSend", _
                    "NdisTransferData", "NdisReset", "NdisRequest", "NdisInitializeWrapper", "NdisTerminateWrapper", "NdisRegisterMac", "NdisDeregisterMac", "NdisRegisterAdapter", "NdisDeregisterAdapter", "NdisCompleteOpenAdapter", "NdisCompleteCloseAdapter", "NdisCompleteSend", "NdisCompleteTransferData", "NdisCompleteReset", "NdisCompleteRequest", _
                    "NdisIndicateReceive", "NdisIndicateReceiveComplete", "NdisIndicateStatus", "NdisIndicateStatusComplete", "NdisCompleteQueryStatistics", "NdisEqualString", "NdisRegAdaptShutdown", "NdisReadNetworkAddress", "NdisWriteErrorLogEntry", "NdisMapIoSpace", "NdisDeregAdaptShutdown", "NdisAllocateSharedMemory", "NdisFreeSharedMemory", "NdisAllocateDmaChannel", "NdisSetupDmaTransfer", _
                    "NdisCompleteDmaTransfer", "NdisReadDmaCounter", "NdisFreeDmaChannel", "NdisReleaseAdapterResources", "NdisQueryGlobalStatistics", "NdisOpenProtocolConfiguration", "NdisCompleteBindAdapter", "NdisCompleteUnbindAdapter", "WrapperStartNet", "WrapperGetComponentList", "WrapperQueryAdapterResources", "WrapperDelayBinding", "WrapperResumeBinding", "WrapperRemoveChildren", "NdisImmediateReadPciSlotInformation", _
                    "NdisImmediateWritePciSlotInformation", "NdisReadPciSlotInformation", "NdisWritePciSlotInformation", "NdisPciAssignResources", "NdisQueryBufferOffset", "NdisMWanSend", "DbgPrint", "NdisInitializeEvent", "NdisSetEvent", "NdisResetEvent", "NdisWaitEvent")

strVMMCall002A = Array( _
                    "VWIN32_Get_Version", "VWIN32_DIOCCompletionRoutine", "_VWIN32_QueueUserApc", "_VWIN32_Get_Thread_Context", "_VWIN32_Set_Thread_Context", "_VWIN32_CopyMem", "_VWIN32_Npx_Exception", "_VWIN32_Emulate_Npx", "_VWIN32_CheckDelayedNpxTrap", "VWIN32_EnterCrstR0", "VWIN32_LeaveCrstR0", "_VWIN32_FaultPopup", "VWIN32_GetContextHandle", "VWIN32_GetCurrentProcessHandle", "_VWIN32_SetWin32Event", "_VWIN32_PulseWin32Event", _
                    "_VWIN32_ResetWin32Event", "_VWIN32_WaitSingleObject", "_VWIN32_WaitMultipleObjects", "_VWIN32_CreateRing0Thread", "_VWIN32_CloseVxDHandle", "VWIN32_ActiveTimeBiasSet", "VWIN32_GetCurrentDirectory", "VWIN32_BlueScreenPopup", "VWIN32_TerminateApp", "_VWIN32_QueueKernelAPC", "VWIN32_SysErrorBox", "_VWIN32_IsClientWin32", "VWIN32_IFSRIPWhenLev2Taken", "_VWIN32_InitWin32Event", "_VWIN32_InitWin32Mutex", _
                    "_VWIN32_ReleaseWin32Mutex", "_VWIN32_BlockThreadEx", "VWIN32_GetProcessHandle", "_VWIN32_InitWin32Semaphore", "_VWIN32_SignalWin32Sem", "_VWIN32_QueueUserApcEx", "_VWIN32_OpenVxDHandle", "_VWIN32_CloseWin32Handle", "_VWIN32_AllocExternalHandle", "_VWIN32_UseExternalHandle", "_VWIN32_UnuseExternalHandle", "KeInitializeTimer", "KeSetTimer", "KeCancelTimer", "KeReadStateTimer", _
                    "_VWIN32_ReferenceObject", "_VWIN32_GetExternalHandle", "VWIN32_ConvertNtTimeout", "_VWIN32_SetWin32EventBoostPriority", "_VWIN32_GetRing3Flat32Selectors", "_VWIN32_GetCurThreadCondition", "VWIN32_Init_FP", "R0SetWaitableTimer")

strVMMCall002B = Array( _
                    "VCOMM_Get_Version", "_VCOMM_Register_Port_Driver", "_VCOMM_Acquire_Port", "_VCOMM_Release_Port", "_VCOMM_OpenComm", "_VCOMM_SetCommState", "_VCOMM_GetCommState", "_VCOMM_SetupComm", "_VCOMM_TransmitCommChar", "_VCOMM_CloseComm", "_VCOMM_GetCommQueueStatus", "_VCOMM_ClearCommError", "_VCOMM_GetModemStatus", "_VCOMM_GetCommProperties", "_VCOMM_EscapeCommFunction", "_VCOMM_PurgeComm", _
                    "_VCOMM_SetCommEventMask", "_VCOMM_GetCommEventMask", "_VCOMM_WriteComm", "_VCOMM_ReadComm", "_VCOMM_EnableCommNotification", "_VCOMM_GetLastError", "_VCOMM_Steal_Port", "_VCOMM_SetReadCallBack", "_VCOMM_SetWriteCallBack", "_VCOMM_Add_Port", "_VCOMM_GetSetCommTimeouts", "_VCOMM_SetWriteRequest", "_VCOMM_SetReadRequest", "_VCOMM_Dequeue_Request", "_VCOMM_Enumerate_DevNodes", _
                    "VCOMM_Map_Win32DCB_To_Ring0", "VCOMM_Map_Ring0DCB_To_Win32", "_VCOMM_Get_Contention_Handler", "_VCOMM_Map_Name_To_Resource", "_VCOMM_PowerOnOffComm")

strVMMCall0033 = Array( _
                    "_CONFIGMG_Get_Version", "_CONFIGMG_Initialize", "_CONFIGMG_Locate_DevNode", "_CONFIGMG_Get_Parent", "_CONFIGMG_Get_Child", "_CONFIGMG_Get_Sibling", "_CONFIGMG_Get_Device_ID_Size", "_CONFIGMG_Get_Device_ID", "_CONFIGMG_Get_Depth", "_CONFIGMG_Get_Private_DWord", "_CONFIGMG_Set_Private_DWord", "_CONFIGMG_Create_DevNode", "_CONFIGMG_Query_Remove_SubTree", "_CONFIGMG_Remove_SubTree", "_CONFIGMG_Register_Device_Driver", "_CONFIGMG_Register_Enumerator", _
                    "_CONFIGMG_Register_Arbitrator", "_CONFIGMG_Deregister_Arbitrator", "_CONFIGMG_Query_Arbitrator_Free_Size", "_CONFIGMG_Query_Arbitrator_Free_Data", "_CONFIGMG_Sort_NodeList", "_CONFIGMG_Yield", "_CONFIGMG_Lock", "_CONFIGMG_Unlock", "_CONFIGMG_Add_Empty_Log_Conf", "_CONFIGMG_Free_Log_Conf", "_CONFIGMG_Get_First_Log_Conf", "_CONFIGMG_Get_Next_Log_Conf", "_CONFIGMG_Add_Res_Des", "_CONFIGMG_Modify_Res_Des", "_CONFIGMG_Free_Res_Des", _
                    "_CONFIGMG_Get_Next_Res_Des", "_CONFIGMG_Get_Performance_Info", "_CONFIGMG_Get_Res_Des_Data_Size", "_CONFIGMG_Get_Res_Des_Data", "_CONFIGMG_Process_Events_Now", "_CONFIGMG_Create_Range_List", "_CONFIGMG_Add_Range", "_CONFIGMG_Delete_Range", "_CONFIGMG_Test_Range_Available", "_CONFIGMG_Dup_Range_List", "_CONFIGMG_Free_Range_List", "_CONFIGMG_Invert_Range_List", "_CONFIGMG_Intersect_Range_List", "_CONFIGMG_First_Range", "_CONFIGMG_Next_Range", _
                    "_CONFIGMG_Dump_Range_List", "_CONFIGMG_Load_DLVxDs", "_CONFIGMG_Get_DDBs", "_CONFIGMG_Get_CRC_CheckSum", "_CONFIGMG_Register_DevLoader", "_CONFIGMG_Reenumerate_DevNode", "_CONFIGMG_Setup_DevNode", "_CONFIGMG_Reset_Children_Marks", "_CONFIGMG_Get_DevNode_Status", "_CONFIGMG_Remove_Unmarked_Children", "_CONFIGMG_ISAPNP_To_CM", "_CONFIGMG_CallBack_Device_Driver", "_CONFIGMG_CallBack_Enumerator", "_CONFIGMG_Get_Alloc_Log_Conf", "_CONFIGMG_Get_DevNode_Key_Size", _
                    "_CONFIGMG_Get_DevNode_Key", "_CONFIGMG_Read_Registry_Value", "_CONFIGMG_Write_Registry_Value", "_CONFIGMG_Disable_DevNode", "_CONFIGMG_Enable_DevNode", "_CONFIGMG_Move_DevNode", "_CONFIGMG_Set_Bus_Info", "_CONFIGMG_Get_Bus_Info", "_CONFIGMG_Set_HW_Prof", "_CONFIGMG_Recompute_HW_Prof", "_CONFIGMG_Query_Change_HW_Prof", "_CONFIGMG_Get_Device_Driver_Private_DWord", "_CONFIGMG_Set_Device_Driver_Private_DWord", "_CONFIGMG_Get_HW_Prof_Flags", "_CONFIGMG_Set_HW_Prof_Flags", _
                    "_CONFIGMG_Read_Registry_Log_Confs", "_CONFIGMG_Run_Detection", "_CONFIGMG_Call_At_Appy_Time", "_CONFIGMG_Fail_Change_HW_Prof", "_CONFIGMG_Set_Private_Problem", "_CONFIGMG_Debug_DevNode", "_CONFIGMG_Get_Hardware_Profile_Info", "_CONFIGMG_Register_Enumerator_Function", "_CONFIGMG_Call_Enumerator_Function", "_CONFIGMG_Add_ID", "_CONFIGMG_Find_Range", "_CONFIGMG_Get_Global_State", "_CONFIGMG_Broadcast_Device_Change_Message", "_CONFIGMG_Call_DevNode_Handler", "_CONFIGMG_Remove_Reinsert_All", _
                    "_CONFIGMG_Change_DevNode_Status", "_CONFIGMG_Reprocess_DevNode", "_CONFIGMG_Assert_Structure", "_CONFIGMG_Discard_Boot_Log_Conf", "_CONFIGMG_Set_Dependent_DevNode", "_CONFIGMG_Get_Dependent_DevNode", "_CONFIGMG_Refilter_DevNode", "_CONFIGMG_Merge_Range_List", "_CONFIGMG_Substract_Range_List", "_CONFIGMG_Set_DevNode_PowerState", "_CONFIGMG_Get_DevNode_PowerState", "_CONFIGMG_Set_DevNode_PowerCapabilities", "_CONFIGMG_Get_DevNode_PowerCapabilities", "_CONFIGMG_Read_Range_List", "_CONFIGMG_Write_Range_List", _
                    "_CONFIGMG_Get_Set_Log_Conf_Priority", "_CONFIGMG_Support_Share_Irq", "_CONFIGMG_Get_Parent_Structure", "_CONFIGMG_Register_DevNode_For_Idle_Detection", "_CONFIGMG_CM_To_ISAPNP", "_CONFIGMG_Get_DevNode_Handler", "_CONFIGMG_Detect_Resource_Conflict", "_CONFIGMG_Get_Device_Interface_List", "_CONFIGMG_Get_Device_Interface_List_Size", "_CONFIGMG_Get_Conflict_Info", "_CONFIGMG_Add_Remove_DevNode_Property", "_CONFIGMG_CallBack_At_Appy_Time", "_CONFIGMG_Register_Device_Interface", "_CONFIGMG_System_Device_Power_State_Mapping", "_CONFIGMG_Get_Arbitrator_Info", _
                    "_CONFIGMG_Waking_Up_From_DevNode", "_CONFIGMG_Set_DevNode_Problem", "_CONFIGMG_Get_Device_Interface_Alias")

strVMMCall0036 = Array( _
                    "VFBACKUP_Get_Version", "VFBACKUP_Lock_NEC", "VFBACKUP_UnLock_NEC", "VFBACKUP_Register_NEC", "VFBACKUP_Register_VFD", "VFBACKUP_Lock_All_Ports", "_VFBACKUP_Set_Port_Mask", "_VFBACKUP_Register_Floppy", "_VFBACKUP_Remove_Floppy", "VFBACKUP_UnRegister_NEC")

strVMMCall0037 = Array( _
                    "VMINI_GetVersion", "VMINI_Update", "VMINI_Status", "VMINI_DisplayError", "VMINI_SetTimeStamp", "VMINI_Siren", "VMINI_RegisterAccess", "VMINI_GetData", "VMINI_ShutDownItem", "VMINI_RegisterSK")

strVMMCall0038 = Array( _
                    "VCOND_Get_Version", "VCOND_Launch_ConApp_Inherited", "VCOND_Get_ConsoleInfo", "VCOND_GrbRepaintRect", "VCOND_GrbSetCursorPosition", "VCOND_GrbNotifyWOA")

strVMMCall0040 = Array( _
                    "IFSMgr_Get_Version", "IFSMgr_RegisterMount", "IFSMgr_RegisterNet", "IFSMgr_RegisterMailSlot", "IFSMgr_Attach", "IFSMgr_Detach", "IFSMgr_Get_NetTime", "IFSMgr_Get_DOSTime", "IFSMgr_SetupConnection", "IFSMgr_DerefConnection", "IFSMgr_ServerDOSCall", "IFSMgr_CompleteAsync", "IFSMgr_RegisterHeap", "IFSMgr_GetHeap", "IFSMgr_RetHeap", "IFSMgr_CheckHeap", _
                    "IFSMgr_CheckHeapItem", "IFSMgr_FillHeapSpare", "IFSMgr_Block", "IFSMgr_Wakeup", "IFSMgr_Yield", "IFSMgr_SchedEvent", "IFSMgr_QueueEvent", "IFSMgr_KillEvent", "IFSMgr_FreeIOReq", "IFSMgr_MakeMailSlot", "IFSMgr_DeleteMailSlot", "IFSMgr_WriteMailSlot", "IFSMgr_PopUp", "IFSMgr_printf", "IFSMgr_AssertFailed", _
                    "IFSMgr_LogEntry", "IFSMgr_DebugMenu", "IFSMgr_DebugVars", "IFSMgr_GetDebugString", "IFSMgr_GetDebugHexNum", "IFSMgr_NetFunction", "IFSMgr_DoDelAllUses", "IFSMgr_SetErrString", "IFSMgr_GetErrString", "IFSMgr_SetReqHook", "IFSMgr_SetPathHook", "IFSMgr_UseAdd", "IFSMgr_UseDel", "IFSMgr_InitUseAdd", "IFSMgr_ChangeDir", _
                    "IFSMgr_DelAllUses", "IFSMgr_CDROM_Attach", "IFSMgr_CDROM_Detach", "IFSMgr_Win32DupHandle", "IFSMgr_Ring0_FileIO", "IFSMgr_Win32_Get_Ring0_Handle", "IFSMgr_Get_Drive_Info", "IFSMgr_Ring0GetDriveInfo", "IFSMgr_BlockNoEvents", "IFSMgr_NetToDosTime", "IFSMgr_DosToNetTime", "IFSMgr_DosToWin32Time", "IFSMgr_Win32ToDosTime", "IFSMgr_NetToWin32Time", "IFSMgr_Win32ToNetTime", _
                    "IFSMgr_MetaMatch", "IFSMgr_TransMatch", "IFSMgr_CallProvider", "UniToBCS", "UniToBCSPath", "BCSToUni", "UniToUpper", "UniCharToOEM", "CreateBasis", "MatchBasisName", "AppendBasisTail", "FcbToShort", "ShortToFcb", "IFSMgr_ParsePath", "Query_PhysLock", _
                    "_VolFlush", "NotifyVolumeArrival", "NotifyVolumeRemoval", "QueryVolumeRemoval", "IFSMgr_FSDUnmountCFSD", "IFSMgr_GetConversionTablePtrs", "IFSMgr_CheckAccessConflict", "IFSMgr_LockFile", "IFSMgr_UnlockFile", "IFSMgr_RemoveLocks", "IFSMgr_CheckLocks", "IFSMgr_CountLocks", "IFSMgr_ReassignLockFileInst", "IFSMgr_UnassignLockList", "IFSMgr_MountChildVolume", _
                    "IFSMgr_UnmountChildVolume", "IFSMgr_SwapDrives", "IFSMgr_FSDMapFHtoIOREQ", "IFSMgr_FSDParsePath", "IFSMgr_FSDAttachSFT", "IFSMgr_GetTimeZoneBias", "IFSMgr_PNPEvent", "IFSMgr_RegisterCFSD", "IFSMgr_Win32MapExtendedHandleToSFT", "IFSMgr_DbgSetFileHandleLimit", "IFSMgr_Win32MapSFTToExtendedHandle", "IFSMgr_FSDGetCurrentDrive", "IFSMgr_InstallFileSystemApiHook", "IFSMgr_RemoveFileSystemApiHook", "IFSMgr_RunScheduledEvents", _
                    "IFSMgr_CheckDelResource", "IFSMgr_Win32GetVMCurdir", "IFSMgr_SetupFailedConnection", "_GetMappedErr", "ShortToLossyFcb", "IFSMgr_GetLockState", "BcsToBcs", "IFSMgr_SetLoopback", "IFSMgr_ClearLoopback", "IFSMgr_ParseOneElement", "BcsToBcsUpper", "IFSMgr_DeregisterFSD", "IFSMgr_RegisterFSDWithPriority", "IFSMgr_Get_DOSTimeRounded", "_LongToFcbOem", _
                    "IFSMgr_GetRing0FileHandle", "IFSMgr_UpdateTimezoneInfo", "IFSMgr_Ring0IsCPSingleByte")

strVMMCall0043 = Array( _
                    "_PCI_Get_Version", "_PCI_Read_Config", "_PCI_Write_Config", "_PCI_Lock_Unlock")

strVMMCall0048 = Array( _
                    "PERF_Get_Version", "PERF_Server_Register", "PERF_Server_Deregister", "PERF_Add_Stat", "PERF_Remove_Stat")

strVMMCall004A = Array( _
                    "_MTRR_Get_Version", "MTRRSetPhysicalCacheTypeRange", "MTRRIsPatSupported", "MTRR_PowerState_Change")

strVMMCall004B = Array( _
                    "_NTKERN_Get_Version", "_NtKernCreateFile", "_NtKernClose", "_NtKernReadFile", "_NtKernWriteFile", "_NtKernDeviceIoControl", "_NtKernGetWorkerThread", "_NtKernLoadDriver", "_NtKernQueueWorkItem", "_NtKernPhisicalDeviceObjectToDevNode", "_NtKernSetPhysicalCahseTypeRange", "_NtKernWin9XLoadDriver", "_NtKernCancelIoFile", "_NtKernGetVPICDHandleFromInterruptObj", "_NtKernInternalDeviceIoControl")

strVMMCalls = Array(, strVMMCall0001, , strVMMCall0003, strVMMCall0004, strVMMCall0005, strVMMCall0006, strVMMCall0007, , , , , strVMMCall000C, strVMMCall000D, strVMMCall000E, , _
                    strVMMCall0010, strVMMCall0011, strVMMCall0012, , strVMMCall0014, strVMMCall0015, , strVMMCall0017, strVMMCall0018, , strVMMCall001A, strVMMCall001B, strVMMCall001C, , , , _
                    strVMMCall0020, strVMMCall0021, , , , , strVMMCall0026, strVMMCall0027, strVMMCall0028, , strVMMCall002A, strVMMCall002B, , , , _
                    , , , , strVMMCall0033, , , strVMMCall0036, strVMMCall0037, strVMMCall0038, , , , , , , , _
                    strVMMCall0040, , , strVMMCall0043, , , , , strVMMCall0048, , strVMMCall004A, strVMMCall004B)
End Sub

Private Sub ProcessEntryTable(ByVal offHeader As Long, header As LE_Header_define)
Dim X As Long, nb As Byte, i As Byte, bDword As Boolean

setPointerOffset offHeader + header.LE_Entry_Table_Offset

X = 0

nb = getByte(0)
Do While nb
    ReDim Preserve arrEntryTable(X)
    
    With arrEntryTable(X)
        .LE_Entry_Number_of_Entries = nb
        .LE_Entry_Bungle_Flags = getByte(0)
        bDword = ((.LE_Entry_Bungle_Flags And LE_EB_32_Bits_Entry) = LE_EB_32_Bits_Entry)
        
        .LE_Entry_Object_Index = getWord(0)
        ReDim .LE_Entry_First_Entry(nb - 1)
        For i = 0 To nb - 1
            With .LE_Entry_First_Entry(i)
                .LE_Entry_Entry_Flags = getByte(0)
                If bDword Then
                    .LE_Entry_Dword_Offset = getDword(0)
                Else
                    .LE_Entry_Dword_Offset = getUWord(0)
                End If
            End With
        Next
    End With
    
    nb = getByte(0)
    X = X + 1
Loop
End Sub

Public Function GetVxDCalls(ByVal dwService As Long) As String
Dim wIdService As Long, wIdIndex As Long

On Error Resume Next
wIdService = (dwService And &HFFFF0000) \ &H10000
wIdIndex = (dwService And &H7FFF&)

GetVxDCalls = strVMMCalls(wIdService)(wIdIndex)
End Function

Private Sub ProcessVXDEntriesName(VxdDesc As VxD_Desc_Block, ServiceTable() As Long)
Dim X As Long, ub As Long
On Error GoTo Fin

ub = UBound(strVMMCalls(VxdDesc.DDB_Req_Device_Number))
If VxdDesc.DDB_Service_Table_Size < (ub + 1) Then ub = VxdDesc.DDB_Service_Table_Size - 1
For X = 0 To ub
    AddName ServiceTable(X), CStr(strVMMCalls(VxdDesc.DDB_Req_Device_Number)(X))
    AddSubName ServiceTable(X), CStr(strVMMCalls(VxdDesc.DDB_Req_Device_Number)(X))
Next

Fin:
End Sub

'Private Sub ProcessVXDCalls(ByVal offHeader As Long, header As LE_Header_define)
'
'End Sub

Private Sub ProcessServiceTable(ByVal offHeader As Long, header As LE_Header_define, VxdDesc As VxD_Desc_Block)
Dim ServT() As Long, X As Long, off As Long

ReDim ServT(VxdDesc.DDB_Service_Table_Size - 1)

AddName VxdDesc.DDB_Service_Table_Ptr, Trim$(StrConv(VxdDesc.DDB_Name, vbUnicode)) & "_Service_Table"
off = VA2Offset(VxdDesc.DDB_Service_Table_Ptr)
getUnkOffset off, VarPtr(ServT(0)), 4 * VxdDesc.DDB_Service_Table_Size
For X = 0 To VxdDesc.DDB_Service_Table_Size - 1
    EntryPointsCol.Add ServT(X)
    setMapOffset off + X * 4, 4
Next
ProcessVXDEntriesName VxdDesc, ServT
End Sub

Private Sub ProcessVXDDesc(ByVal offHeader As Long, header As LE_Header_define, ByVal iFileLE As Integer)
Dim off As Long, va As Long
Dim VD As VxD_Desc_Block

off = retSectionTables(arrEntryTable(0).LE_Entry_Object_Index - 1).PointerToRawData + arrEntryTable(0).LE_Entry_First_Entry(0).LE_Entry_Dword_Offset
va = Offset2VA(off)
AddName va, strVXDDesc
getUnkOffset off, VarPtr(VD), Len(VD)

setMapOffset off, 32
setMapOffset off + 4, 31
setMapOffset off + 6, 31
setMapOffset off + 8, 30
setMapOffset off + 9, 30
setMapOffset off + 10, 31
setMapOffset off + 12, 3
setMapOffset off + 20, 32
setMapOffset off + 24, 4
setMapOffset off + 28, 32
setMapOffset off + 32, 32
setMapOffset off + 36, 32
setMapOffset off + 40, 32
setMapOffset off + 44, 32
setMapOffset off + 48, 4
setMapOffset off + 52, 32
setMapOffset off + 56, 32
setMapOffset off + 60, 32
setMapOffset off + 64, 32

AddName va + 4, "DDB_SDK_Version"
AddName va + 6, "DDB_Req_Device_Number"
AddName va + 8, "DDB_Dev_Major_Version"
AddName va + 9, "DDB_Dev_Minor_Version"
AddName va + 10, "DDB_Flags"
AddName va + 12, "DDB_Name"
AddName va + 20, "DDB_Init_Order"
AddName va + 24, "DDB_Control_Proc"
AddName va + 28, "DDB_V86_API_Proc"
AddName va + 32, "DDB_PM_API_Proc"
AddName va + 36, "DDB_V86_API_CSIP"
AddName va + 40, "DDB_PM_API_CSIP"
AddName va + 44, "DDB_Reference_Data"
AddName va + 48, "DDB_Service_Table_Ptr"
AddName va + 52, "DDB_Service_Table_Size"
AddName va + 56, "DDB_Win32_Service_Table"
AddName va + 60, "DDB_Prev_0"
AddName va + 64, "DDB_Size_0"

Print #iFileLE, "======================================================================"
Print #iFileLE, "VxD Descriptor Block "
Print #iFileLE, "======================================================================"
With VD
    Print #iFileLE, "DDB_Name :", StrConv(.DDB_Name, vbUnicode)
    Print #iFileLE, "DDB_Dev_Major_Version :", .DDB_Dev_Major_Version
    Print #iFileLE, "DDB_Dev_Minor_Version :", .DDB_Dev_Minor_Version
    
    Print #iFileLE, "DDB_Req_Device_Number :", .DDB_Req_Device_Number
    
    Print #iFileLE, "DDB_Flags :", .DDB_Flags
    Print #iFileLE, "DDB_Init_Order :", .DDB_Init_Order
    Print #iFileLE, "DDB_SDK_Version :", .DDB_SDK_Version
    
    Print #iFileLE, "DDB_Control_Proc :", .DDB_Control_Proc
    Print #iFileLE, "DDB_Service_Table_Ptr :", .DDB_Service_Table_Ptr
    Print #iFileLE, "DDB_Service_Table_Size :", .DDB_Service_Table_Size
    
    Print #iFileLE, "DDB_Win32_Service_Table :", .DDB_Win32_Service_Table
    
    Print #iFileLE, "DDB_Reference_Data :", .DDB_Reference_Data
    
    Print #iFileLE, "DDB_PM_API_CSIP :", .DDB_PM_API_CSIP
    Print #iFileLE, "DDB_PM_API_Proc :", .DDB_PM_API_Proc
    Print #iFileLE, "DDB_V86_API_CSIP :", .DDB_V86_API_CSIP
    Print #iFileLE, "DDB_V86_API_Proc :", .DDB_V86_API_Proc

    AddName .DDB_Control_Proc, Trim$(StrConv(.DDB_Name, vbUnicode)) & "_Control"
    AddSubName .DDB_Control_Proc, Trim$(StrConv(.DDB_Name, vbUnicode)) & "_Control"
    EntryPointsCol.Add .DDB_Control_Proc
'    .DDB_Next
'    .DDB_Dev_Major_Version
'    .DDB_Dev_Minor_Version
'    .DDB_Flags
'    .DDB_Init_Order
'    .DDB_Name
'    .DDB_V86_API_CSIP
'    .DDB_V86_API_Proc
'    .DDB_PM_API_CSIP
'    .DDB_PM_API_Proc
'    .DDB_Reference_Data
'    .DDB_SDK_Version

'    .DDB_Req_Device_Number
'    .DDB_Service_Table_Ptr
'    .DDB_Service_Table_Size
    ProcessServiceTable offHeader, header, VD
End With
End Sub

Private Sub ProcessNonResidentNameTable(ByVal offHeader As Long, header As LE_Header_define)
Dim cb As Byte, i As Byte, strName As String, c As Byte, wOrdinal As Integer
Dim oldOffset As Long

If (header.LE_Nonresident_Names_Table_Offset > 0) And (header.LE_Nonresident_Names_Table_Length > 0) Then
    setPointerOffset header.LE_Nonresident_Names_Table_Offset
    
    'module name
    cb = getByte(0)
    strModuleDescription = vbNullString
    For i = 1 To cb
        c = getByte(0)
        strModuleDescription = strModuleDescription & Chr$(c)
    Next
    wOrdinal = getWord(0)
    
    'VXD Description Block
    cb = getByte(0)
    strVXDDesc = vbNullString
    For i = 1 To cb
        c = getByte(0)
        strVXDDesc = strVXDDesc & Chr$(c)
    Next
    wOrdinal = getWord(0)
    
    cb = getByte(0)
    Do While cb
        strName = vbNullString
        For i = 1 To cb
            c = getByte(0)
            strName = strName & Chr$(c)
        Next
        wOrdinal = getWord(0)
        
        Debug.Assert False
        
        cb = getByte(0)
    Loop
End If
End Sub
'
'Private Sub ProcessImportedModulesNameTable()
''
'End Sub
'
'Private Sub ProcessImportedProcsNameTable()
''
'End Sub

Private Sub ProcessFixupPageTable(ByVal offHeader As Long, header As LE_Header_define)
Dim X As Long, offFixupCurrentPage As Long, offFixupNextPage As Long, dwRelocSize As Long
Dim fixup As LE_Fixup_Record_Table_Define, i As Integer
Dim off As Long, reloc As Long

For X = 0 To header.LE_Number_Of_Memory_Pages - 1
    offFixupCurrentPage = getDwordOffset(offHeader + header.LE_Fixup_Page_Table_Offset + X * 4)
    offFixupNextPage = getDwordOffset(offHeader + header.LE_Fixup_Page_Table_Offset + (X + 1) * 4)
    
    dwRelocSize = offFixupNextPage - offFixupCurrentPage
    If dwRelocSize > 0 Then
        setPointerOffset offHeader + header.LE_Fixup_Record_Table_Offset + offFixupCurrentPage
        
        Do While dwRelocSize > 0
            With fixup
'                If getPointerOffset >= &H3000& Then
'                    Debug.Assert False
'                End If
                .LE_Fixup_Relocation_Address_Type = getByte(0)
                .LE_Fixup_Relocation_Type = getByte(0)
                dwRelocSize = dwRelocSize - 2
                
                If (.LE_Fixup_Relocation_Address_Type And LE_RAT_List_Offset) = 0 Then
                    .LE_Fixup_Relocation_Page_Offset = getUWord(0)
                    dwRelocSize = dwRelocSize - 2
                    Select Case (.LE_Fixup_Relocation_Type And LE_RT_Reloc_Type)
                        Case LE_RT_Internal_Reference
                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
                                .LE_Fixup_Offset_Or_Ordinal_Value = getDword(0)
                                dwRelocSize = dwRelocSize - 5
                            Else
                                .LE_Fixup_Offset_Or_Ordinal_Value = getUWord(0)
                                dwRelocSize = dwRelocSize - 3
                            End If
                            
                            Select Case (.LE_Fixup_Relocation_Address_Type And LE_RAT_Rel_Addr_Type)
                                Case LE_RA_32_bits_EIP_Rel
                                    off = header.LE_First_Pages_Offset + X * header.LE_Memory_Page_Size + .LE_Fixup_Relocation_Page_Offset
                                    reloc = (retSectionTables(.LE_Fixup_Segment_or_Module_Index - 1).PointerToRawData + .LE_Fixup_Offset_Or_Ordinal_Value) - off - 4
                                    setDwordOffset off, reloc
                                Case LE_RA_32_bits_Offset
                                    off = header.LE_First_Pages_Offset + X * header.LE_Memory_Page_Size + .LE_Fixup_Relocation_Page_Offset
                                    reloc = retSectionTables(.LE_Fixup_Segment_or_Module_Index - 1).VirtualAddress + .LE_Fixup_Offset_Or_Ordinal_Value
                                    setDwordOffset off, reloc
                                'Case LE_RA_16_bits_Offset
                                'Case LE_RA_16_bits_selector
                                'Case LE_RA_32_bits_Far_Pointer
                                'Case LE_RA_48_bits_Far_Pointer
                                'Case LE_RA_Low_Byte '???
                            End Select
                        Case Else
                            Debug.Assert False
'                        Case LE_RT_Imported_Ordinal
'                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
'                            dwRelocSize = dwRelocSize - 1
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Ordinal_Byte) = LE_RT_Ordinal_Byte Then
'                                .LE_Fixup_Offset_Or_Ordinal_Value = getByte(0)
'                                dwRelocSize = dwRelocSize - 1
'                            Else
'                                .LE_Fixup_Offset_Or_Ordinal_Value = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_ADDITIVE_Type) = LE_RT_ADDITIVE_Type Then
'                                If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                    .LE_Fixup_AddValue_ = getDword(0)
'                                    dwRelocSize = dwRelocSize - 4
'                                Else
'                                    .LE_Fixup_AddValue_ = getWord(0)
'                                    dwRelocSize = dwRelocSize - 2
'                                End If
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                .LE_Fixup_Extra_ = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'
'                            Select Case (.LE_Fixup_Relocation_Address_Type And LE_RAT_Rel_Addr_Type)
'                                Case LE_RA_16_bits_Offset
'                                Case LE_RA_16_bits_selector
'                                Case LE_RA_32_bits_EIP_Rel
'                                Case LE_RA_32_bits_Far_Pointer
'                                Case LE_RA_32_bits_Offset
'                                Case LE_RA_32_bits_Offset
'                                Case LE_RA_48_bits_Far_Pointer
'                                Case LE_RA_Low_Byte
'                            End Select
'                        Case LE_RT_Imported_Name
'                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
'                            .LE_Fixup_Offset_Or_Ordinal_Value = getWord(0)
'                            dwRelocSize = dwRelocSize - 3
'                            If (.LE_Fixup_Relocation_Type And LE_RT_ADDITIVE_Type) = LE_RT_ADDITIVE_Type Then
'                                If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                    .LE_Fixup_AddValue_ = getDword(0)
'                                    dwRelocSize = dwRelocSize - 4
'                                Else
'                                    .LE_Fixup_AddValue_ = getWord(0)
'                                    dwRelocSize = dwRelocSize - 2
'                                End If
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                .LE_Fixup_Extra_ = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'
'                            Select Case (.LE_Fixup_Relocation_Address_Type And LE_RAT_Rel_Addr_Type)
'                                Case LE_RA_16_bits_Offset
'                                Case LE_RA_16_bits_selector
'                                Case LE_RA_32_bits_EIP_Rel
'                                Case LE_RA_32_bits_Far_Pointer
'                                Case LE_RA_32_bits_Offset
'                                Case LE_RA_32_bits_Offset
'                                Case LE_RA_48_bits_Far_Pointer
'                                Case LE_RA_Low_Byte
'                            End Select
                    End Select
                Else
                    .LE_Fixup_Offset_Counter = getByte(0)
                    dwRelocSize = dwRelocSize - 1
                    Select Case (.LE_Fixup_Relocation_Type And LE_RT_Reloc_Type)
                        Case LE_RT_Internal_Reference
                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
                            dwRelocSize = dwRelocSize - 1
                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
                                .LE_Fixup_Offset_Or_Ordinal_Value = getDword(0)
                                dwRelocSize = dwRelocSize - 4
                            Else
                                .LE_Fixup_Offset_Or_Ordinal_Value = getUWord(0)
                                dwRelocSize = dwRelocSize - 2
                            End If
                            
                            dwRelocSize = dwRelocSize - 2 * .LE_Fixup_Offset_Counter
                            For i = 0 To .LE_Fixup_Offset_Counter - 1
                                .LE_Fixup_Relocation_Page_Offset = getUWord(0)
                                Select Case (.LE_Fixup_Relocation_Address_Type And LE_RAT_Rel_Addr_Type)
                                    Case LE_RA_32_bits_EIP_Rel
                                        off = header.LE_First_Pages_Offset + X * header.LE_Memory_Page_Size + .LE_Fixup_Relocation_Page_Offset
                                        reloc = (retSectionTables(.LE_Fixup_Segment_or_Module_Index - 1).PointerToRawData + .LE_Fixup_Offset_Or_Ordinal_Value) - off - 4
                                        setDwordOffset off, reloc
                                    Case LE_RA_32_bits_Offset
                                        off = header.LE_First_Pages_Offset + X * header.LE_Memory_Page_Size + .LE_Fixup_Relocation_Page_Offset
                                        reloc = retSectionTables(.LE_Fixup_Segment_or_Module_Index - 1).VirtualAddress + .LE_Fixup_Offset_Or_Ordinal_Value
                                        setDwordOffset off, reloc
                                    'Case LE_RA_16_bits_Offset
                                    'Case LE_RA_16_bits_selector
                                    'Case LE_RA_32_bits_Far_Pointer
                                    'Case LE_RA_48_bits_Far_Pointer
                                    'Case LE_RA_Low_Byte '???
                                End Select
                            Next
                        Case Else
                            Debug.Assert False
'                        Case LE_RT_Imported_Ordinal
'                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
'                            dwRelocSize = dwRelocSize - 1
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Ordinal_Byte) = LE_RT_Ordinal_Byte Then
'                                .LE_Fixup_Offset_Or_Ordinal_Value = getByte(0)
'                                dwRelocSize = dwRelocSize - 1
'                            Else
'                                .LE_Fixup_Offset_Or_Ordinal_Value = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_ADDITIVE_Type) = LE_RT_ADDITIVE_Type Then
'                                If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                    .LE_Fixup_AddValue_ = getDword(0)
'                                    dwRelocSize = dwRelocSize - 4
'                                Else
'                                    .LE_Fixup_AddValue_ = getWord(0)
'                                    dwRelocSize = dwRelocSize - 2
'                                End If
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                .LE_Fixup_Extra_ = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'                            dwRelocSize = dwRelocSize - 2 * .LE_Fixup_Relocation_Page_Offset_Or_Offset_Counter
'                            For i = 0 To .LE_Fixup_Relocation_Page_Offset_Or_Offset_Counter - 1
'                                .LE_Fixup_Offset_ = getWord(0)
'
'
'                            Next
'                        Case LE_RT_Imported_Name
'                            .LE_Fixup_Segment_or_Module_Index = getByte(0)
'                            .LE_Fixup_Offset_Or_Ordinal_Value = getWord(0)
'                            dwRelocSize = dwRelocSize - 3
'                            If (.LE_Fixup_Relocation_Type And LE_RT_ADDITIVE_Type) = LE_RT_ADDITIVE_Type Then
'                                If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                    .LE_Fixup_AddValue_ = getDword(0)
'                                    dwRelocSize = dwRelocSize - 4
'                                Else
'                                    .LE_Fixup_AddValue_ = getWord(0)
'                                    dwRelocSize = dwRelocSize - 2
'                                End If
'                            End If
'                            If (.LE_Fixup_Relocation_Type And LE_RT_Target_Offset_32) = LE_RT_Target_Offset_32 Then
'                                .LE_Fixup_Extra_ = getWord(0)
'                                dwRelocSize = dwRelocSize - 2
'                            End If
'                            dwRelocSize = dwRelocSize - 2 * .LE_Fixup_Relocation_Page_Offset_Or_Offset_Counter
'                            For i = 0 To .LE_Fixup_Relocation_Page_Offset_Or_Offset_Counter - 1
'                                .LE_Fixup_Offset_ = getWord(0)
'
'
'                            Next
                    End Select
                End If
            End With
        Loop
    End If
Next
End Sub

Private Sub ProcessResidentNameTable(ByVal offHeader As Long, header As LE_Header_define)
Dim cb As Byte, i As Byte, strName As String, c As Byte, wOrdinal As Integer

If header.LE_Resident_Names_Table_Offset Then
    setPointerOffset offHeader + header.LE_Resident_Names_Table_Offset
    
    'module name
    cb = getByte(0)
    strModuleName = vbNullString
    For i = 1 To cb
        c = getByte(0)
        strModuleName = strModuleName & Chr$(c)
    Next
    wOrdinal = getUWord(0)
    
    cb = getByte(0)
    Do While cb
        strName = vbNullString
        For i = 1 To cb
            c = getByte(0)
            strName = strName & Chr$(c)
        Next
        wOrdinal = getUWord(0)
        
        Debug.Assert False
        
        cb = getByte(0)
    Loop
End If
End Sub

Private Sub ProcessObjectTable(ByVal offHeader As Long, header As LE_Header_define, ByVal iFileLE As Integer)
Dim obj As LE_Object_Table_Define, X As Long, off As Long
Dim objmap As LE_Page_Map_Table_Define

'ici modif
'With frmProgress.pbSection1
'    .Min = 0
'    .Max = header.LE_Object_Table_Entries - 1
'End With

Print #iFileLE, "======================================================================"
Print #iFileLE, "Objects Table "
Print #iFileLE, "======================================================================"
With header
    ReDim retSectionTables(.LE_Object_Table_Entries - 1)
    For X = 0 To .LE_Object_Table_Entries - 1
        getUnkOffset offHeader + .LE_Object_Table_Offset + X * 24, VarPtr(obj), Len(obj)
        getUnkOffset offHeader + .LE_Object_Page_Map_Table_Offset + (obj.LE_OBJ_Page_MAP_Index - 1) * 4, VarPtr(objmap), Len(objmap)
        With retSectionTables(X)
            .PointerToRawData = header.LE_First_Pages_Offset + (objmap.LE_PM_High_Page_Number + objmap.LE_PM_Low_Page_Number - 1) * header.LE_Memory_Page_Size
            .SecName(0) = obj.LE_OBJ_Name(0)
            .SecName(1) = obj.LE_OBJ_Name(1)
            .SecName(2) = obj.LE_OBJ_Name(2)
            .SecName(3) = obj.LE_OBJ_Name(3)
            .SecName(4) = 0
            .SecName(5) = 0
            .SecName(6) = 0
            .SecName(7) = 0
            .SizeOfRawData = obj.LE_OBJ_Page_MAP_Entries * header.LE_Memory_Page_Size
            .VirtualAddress = dwImageBase + (objmap.LE_PM_High_Page_Number + objmap.LE_PM_Low_Page_Number - 1) * header.LE_Memory_Page_Size
            .VirtualSize = obj.LE_OBJ_Virtual_Segment_Size
            

            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_BIG_Segment) = 0 Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_16BIT
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_16_16_Alias) = LE_OBJ_FL_16_16_Alias Then .Characteristics = .Characteristics Or image_scn_
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Conforming_Segment) = LE_OBJ_FL_Conforming_Segment Then .Characteristics = .Characteristics Or image_scn_
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_I_O_Privilage_Level) = LE_OBJ_FL_I_O_Privilage_Level Then .Characteristics = .Characteristics Or image_scn_
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Reserved) = LE_OBJ_FL_Reserved Then .Characteristics = .Characteristics Or image_scn_
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Resident_Long_Locable) = LE_OBJ_FL_Resident_Long_Locable Then .Characteristics = .Characteristics Or image_scn_
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Discardable) = LE_OBJ_FL_Segment_Discardable Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_DISCARDABLE
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Executable) = LE_OBJ_FL_Segment_Executable Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_EXECUTE Or IMAGE_SCN_CNT_CODE
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Invalid) = LE_OBJ_FL_Segment_Invalid Then .Characteristics = .Characteristics Or IMAGE_SCN_CNT_UNINITIALIZED_DATA
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Preloaded) = LE_OBJ_FL_Segment_Preloaded Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_PRELOAD
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Readable) = LE_OBJ_FL_Segment_Readable Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_READ
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Resident) = LE_OBJ_FL_Segment_Resident Then .Characteristics = .Characteristics Or image_scn_
            'If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Resident_contiguous) = LE_OBJ_FL_Segment_Resident_contiguous Then .Characteristics = .Characteristics Or image_scn_
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Resource) = LE_OBJ_FL_Segment_Resource Then .Characteristics = .Characteristics Or IMAGE_SCN_CNT_INITIALIZED_DATA
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Shared) = LE_OBJ_FL_Segment_Shared Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_SHARED
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Writable) = LE_OBJ_FL_Segment_Writable Then .Characteristics = .Characteristics Or IMAGE_SCN_MEM_WRITE
            If (obj.LE_OBJ_FLAGS And LE_OBJ_FL_Segment_Zero_Filled) = LE_OBJ_FL_Segment_Zero_Filled Then .Characteristics = .Characteristics Or IMAGE_SCN_CNT_UNINITIALIZED_DATA
            
            For off = .PointerToRawData + .VirtualSize To .PointerToRawData + .SizeOfRawData - 1
                setMapOffset off, 255
            Next
        End With
        Print #iFileLE, "----------------------------------------------------------------------"
        Print #iFileLE, "Object ", X + 1
        Print #iFileLE, "----------------------------------------------------------------------"
        With obj
            Print #iFileLE, "Name :", StrConv(.LE_OBJ_Name, vbUnicode)
            Print #iFileLE, "Relocation Base Address :", .LE_OBJ_Relocation_Base_Address
            Print #iFileLE, "Virtual Segment Size :", .LE_OBJ_Virtual_Segment_Size
            Print #iFileLE, "FLAGS :", .LE_OBJ_FLAGS
            Print #iFileLE, "Page MAP Entries :", .LE_OBJ_Page_MAP_Entries
            Print #iFileLE, "Page MAP Index :", .LE_OBJ_Page_MAP_Index
        End With
        'frmProgress.pbSection1.value = X
        DoEvents
    Next
End With
Print #iFileLE, "======================================================================"
End Sub

Private Sub ProcessLEFile(szFilename As String, szOutPattern As String)
Dim Offset As Long, iFileLE As Integer
Dim hd As LE_Header_define

iFileLE = FreeFile
Open szOutPattern & ".le" For Output As #iFileLE
    Print #iFileLE, "======================================================================"
    Print #iFileLE, "VxD (LE) File : "; szFilename
    Print #iFileLE, "======================================================================"

    Offset = getDwordOffset(&H3C)
    getUnkOffset Offset, VarPtr(hd), Len(hd)
    
    ProcessResidentNameTable Offset, hd
    
    ProcessNonResidentNameTable Offset, hd
    
    With hd
        Print #iFileLE, "Module Name :", strModuleName
        Print #iFileLE, "Module Description :", strModuleDescription
        Print #iFileLE, "VXD Descriptor Name :", strVXDDesc
        
        Print #iFileLE, "LE_CPU_Type :", .LE_CPU_Type
        Print #iFileLE, "LE_Byte_Order :", .LE_Byte_Order
        Print #iFileLE, "LE_Word_Order :", .LE_Word_Order
        Print #iFileLE, "LE_Target_OS :", .LE_Target_OS
        Print #iFileLE, "LE_Module_Type_Flags :", .LE_Module_Type_Flags
        Print #iFileLE, "LE_Module_Version :", .LE_Module_Version
        Print #iFileLE, "LE_Exec_Format_Level :", .LE_Exec_Format_Level
        
        Print #iFileLE, "LE_Initial_CS :", .LE_Initial_CS
        Print #iFileLE, "LE_Initial_EIP :", .LE_Initial_EIP
        Print #iFileLE, "LE_Initial_SS :", .LE_Initial_SS
        Print #iFileLE, "LE_Initial_ESP :", .LE_Initial_ESP
        
        Print #iFileLE, "LE_First_Pages_Offset :", .LE_First_Pages_Offset
        Print #iFileLE, "LE_Memory_Page_Size :", .LE_Memory_Page_Size
        Print #iFileLE, "LE_Number_Of_Memory_Pages :", .LE_Number_Of_Memory_Pages
        Print #iFileLE, "LE_Bytes_On_Last_Page :", .LE_Bytes_On_Last_Page
        
        Print #iFileLE, "LE_Extra_Heap_Allocation :", .LE_Extra_Heap_Allocation
        
        Print #iFileLE, "LE_Entry_Table_Offset :", .LE_Entry_Table_Offset
        
        Print #iFileLE, "LE_Fixup_Page_Table_Offset :", .LE_Fixup_Page_Table_Offset
        Print #iFileLE, "LE_Fixup_Record_Table_Offset :", .LE_Fixup_Record_Table_Offset
        Print #iFileLE, "LE_Fixup_Section_Checksum :", .LE_Fixup_Section_Checksum
        Print #iFileLE, "LE_Fixup_Section_Size :", .LE_Fixup_Section_Size
        
        Print #iFileLE, "LE_Imported_Module_Names_Table_Offset :", .LE_Imported_Module_Names_Table_Offset
        Print #iFileLE, "LE_Imported_Modules_Count :", .LE_Imported_Modules_Count
        Print #iFileLE, "LE_Imported_Procedure_Name_Table_Offset :", .LE_Imported_Procedure_Name_Table_Offset
        
        Print #iFileLE, "LE_Nonresident_Names_Table_Checksum :", .LE_Nonresident_Names_Table_Checksum
        Print #iFileLE, "LE_Nonresident_Names_Table_Length :", .LE_Nonresident_Names_Table_Length
        Print #iFileLE, "LE_Nonresident_Names_Table_Offset :", .LE_Nonresident_Names_Table_Offset
        
        Print #iFileLE, "LE_Object_Table_Offset :", .LE_Object_Table_Offset
        Print #iFileLE, "LE_Object_Table_Entries :", .LE_Object_Table_Entries
        Print #iFileLE, "LE_Object_Page_Map_Table_Offset :", .LE_Object_Page_Map_Table_Offset
        
        Print #iFileLE, "LE_Resident_Names_Table_Offset :", .LE_Resident_Names_Table_Offset
        
        Print #iFileLE, "LE_Loader_Section_CheckSum :", .LE_Loader_Section_CheckSum
        Print #iFileLE, "LE_Loader_Section_Size :", .LE_Loader_Section_Size
        
        Print #iFileLE, "LE_Resource_Table_Entries :", .LE_Resource_Table_Entries
        Print #iFileLE, "LE_Resource_Table_Offset :", .LE_Resource_Table_Offset
    
        Print #iFileLE, "LE_Debug_Information_Length :", .LE_Debug_Information_Length
        Print #iFileLE, "LE_Debug_Information_Offset :", .LE_Debug_Information_Offset
        
        Print #iFileLE, "LE_Module_Directives_Table_Entries :", .LE_Module_Directives_Table_Entries
        Print #iFileLE, "LE_Module_Directives_Table_Offset :", .LE_Module_Directives_Table_Offset
    
        Print #iFileLE, "LE_Automatic_Data_Object :", .LE_Automatic_Data_Object
        Print #iFileLE, "LE_Demand_Instance_Pages_Number :", .LE_Demand_Instance_Pages_Number
        Print #iFileLE, "LE_Object_Iterate_Data_Map_Offset :", .LE_Object_Iterate_Data_Map_Offset
        Print #iFileLE, "LE_Per_page_Checksum_Table_Offset :", .LE_Per_page_Checksum_Table_Offset
        Print #iFileLE, "LE_Preload_Instance_Pages_Number :", .LE_Preload_Instance_Pages_Number
        Print #iFileLE, "LE_Preload_Page_Count :", .LE_Preload_Page_Count
    End With
    
    ProcessObjectTable Offset, hd, iFileLE
    
    'frmProgress.imSection.Visible = True
    DoEvents
    
    ProcessEntryTable Offset, hd
    
    ProcessFixupPageTable Offset, hd
        
    ProcessVXDDesc Offset, hd, iFileLE
        
    If hd.LE_Initial_CS Then
        InitEntryPoint = retSectionTables(hd.LE_Initial_CS - 1).VirtualAddress + hd.LE_Initial_EIP
    End If
Close #iFileLE
End Sub

Public Sub DysLEFile(szLEFile As String, szOutPattern As String)
Dim iFileASM As Integer, X As Long, addr As Long, b As Byte, dw As Long
Dim iFileDAT As Integer, iFileLOG As Integer

'Load frmProgress
'frmProgress.InitLE
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & szLEFile
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init
InitVMMCalls
Set EntryPointsCol = New Collection

dwImageBase = &H1000000
'chargement du fichier
If LoadFile2(szLEFile) = 0 Then Exit Sub

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents

ProcessLEFile szLEFile, szOutPattern

iFileASM = FreeFile
Open szOutPattern & ".asm" For Output As #iFileASM
iFileDAT = FreeFile
Open szOutPattern & ".dat" For Output As #iFileDAT
iFileLOG = FreeFile
Open szOutPattern & ".log" For Output As #iFileLOG

    'frmProgress.lblState.Caption = "Traitement du point d'entrée..."
    DoEvents
    If InitEntryPoint Then
        If IsCode16VA(InitEntryPoint) Then
            Set16BitsDecode
        Else
            Set32BitsDecode
        End If
        DysCode iFileASM, InitEntryPoint, True, "start"
    End If
    'frmProgress.imStart.Visible = True
    DoEvents
    
    If EntryPointsCol.Count Then
        'frmProgress.lblState.Caption = "Traitement des services..."
        DoEvents
       ' With frmProgress.pbExp
         '   .Min = 0
         '   .Max = EntryPointsCol.Count
            For X = 1 To EntryPointsCol.Count
                addr = EntryPointsCol(X)
                If IsCode16VA(addr) Then
                    Set16BitsDecode
                Else
                    Set32BitsDecode
                End If
                DysCode iFileASM, addr, True, GetSubName(addr)
              '  .value = X
            Next
       ' End With
    End If
    'frmProgress.imExp.Visible = True
    DoEvents
    
   ' frmProgress.lblState.Caption = "Traitement des offsets..."
    DoEvents
    'With frmProgress.pbTry
        '.Min = 1
        X = 1
        Do While X <= tryCallCol.Count
            '.Max = tryCallCol.Count
            
            addr = tryCallCol(X)
            b = getMapVA(addr)
            If b = 0 Then
                dw = GetAddrSize(addr)
                If dw = 1 Then
                    setMapVA addr, 30
                ElseIf dw = 2 Then
                    setMapVA addr, 31
                ElseIf dw = 4 Then
                    setMapVA addr, 32
                Else
                    If IsValidUnicodeString(addr) Then
                        setMapVA addr, 10
                    ElseIf IsValidNullString(addr) Then
                        setMapVA addr, 5
                    ElseIf IsValidPascalString(addr) Then
                        setMapVA addr, 7
                    Else 'numérique
                        dw = getDwordVA(addr)
                        If CheckVA(dw) Then
                            'pointeur
                            setMapVA addr, 4
                            ProcessPointer addr
                        ElseIf dw Then
                            'code
                            dw = getDwordVA(addr)
                            If dw = 0 Then
                                setMapVA addr, 3
                            Else
                                setMapVA addr, 0
                                DysCode iFileASM, addr, True, GetSubName(addr)
                                Print #iFileLOG, "Disassembling from offset at :", getNumber(addr, 8)
                            End If
                        Else
                            setMapVA addr, 3
                        End If
                    End If
                End If
            End If
            '.value = X
            DoEvents
            X = X + 1
        Loop
    'End With
    'frmProgress.imOff.Visible = True
    DoEvents

    'frmProgress.lblState.Caption = "Traitement des données..."
    DoEvents
    
    ProcessData iFileASM, iFileDAT, iFileLOG
    
    'frmProgress.imData.Visible = True
    DoEvents
Close #iFileLOG
Close #iFileDAT
Close #iFileASM

'frmProgress.lblState.Caption = "File disassembled in " & Format$(StopTimer, "#.##") & " seconds"

UnloadFile2
End Sub

