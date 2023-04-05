#  this is the object define for parameter loaded

# import for excel control
from datetime import datetime
import xlwings as xw
# this import is for the VBA function
import win32com.client
# application of array
import numpy as np
# include for atof function => transfer string to float
import locale as lo

# # also for the jump out window, same group with win32con
import win32api

import time
# import for the program exit sys.exit()
import sys


# ======== excel application related
# 開啟 Excel 的app
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# ======== excel application related


class excel_parameter ():
    def __init__(self, book_name):
        # when define the object, open the initialization and load all the parameter
        # into the object with related book name
        # this can help to change the setting by giving different book_name to
        # the verification oject
        # each verification object can get different parameter from different
        # excel_parameter object input
        print('start of the parameter loaded')

        self.control_book_trace = 'c:\\py_gary\\test_excel\\' + \
            str(book_name) + '.xlsm'

        # slection for the control parameter mapped to the object
        if book_name != 'obj_main':
            # open the new book for loading the control parameter in related object
            self.wb = xw.Book(self.control_book_trace)
            print('select other book')
            # 220901 first not used the function to switch to other books

            pass
        else:
            # if book loaded is the original control book, not going to open the book and
            # use the obj_main as the control parameter input
            self.wb = xw.books('obj_main.xlsm')
            print('select original book')

            pass

        # 220907 add the sheet array for eff measurement
        self.sheet_arry = np.full([200], None)

        # after choosing the workbook, define the main sheet to load parameter
        self.sh_main = self.wb.sheets('main')

        # only the instrument control will be still mapped to the original excel
        # since inst_ctrl is no needed to copy to the result sheet
        self.sh_inst_ctrl = self.wb.sheets('inst_ctrl')

        # other way to define sheet:
        # this is the format for the efficiency result
        ex_sheet_name = 'raw_out'
        self.sh_raw_out = self.wb.sheets(ex_sheet_name)
        # this is the sheet for efficiency testing command
        self.sh_volt_curr_cmd = self.wb.sheets('V_I_com')
        # this is the sheet for I2C command
        self.sh_i2c_cmd = self.wb.sheets('I2C_ctrl')
        # this is the sheet for IQ scan
        self.sh_iq_scan = self.wb.sheets('IQ_measured')
        # this is the sheet for wire scan
        self.sh_sw_scan = self.wb.sheets('SWIRE_scan')
        # the sheet used to generate general format of waveform
        self.sh_format_gen = self.wb.sheets('CTRL_sh_ex_ripple')
        # the sheet used save the file information
        self.sh_ref_table = self.wb.sheets('table')
        # thesheet for general testing item
        self.sh_general_test = self.wb.sheets('general_example')

        # file name from the master excel
        self.new_file_name = str(self.sh_main.range('B8').value)
        # load the trace as string
        self.excel_temp = str(self.sh_main.range('B9').value)
        # program control variable (auto and inst settings)
        self.auto_inst_ctrl = self.sh_main.range('B11').value
        # program exit interrupt variable
        self.program_exit = self.sh_main.range('B12').value
        # add for the index of instrument closed or not 1 => inst are close and ready to leave program
        self.turn_inst_off = 0
        self.ready_to_off = 0
        # verification re-run
        self.re_run_verification = self.sh_main.range('B13').value
        # file settings; decide which file to load parameters
        self.file_setting = self.sh_main.range('B14').value
        # program selection setting
        self.program_group_index = self.sh_main.range('B15').value

        # Vin status global variable
        self.vin_status = ''
        # I_AVDD status global variable
        self.i_avdd_status = ''
        # I_EL staatus global variable
        self.i_el_status = ''
        # SW_I2C status global variable
        self.sw_i2c_status = ''

        # default extra file name is call _temp
        # only multi item case will have temp file
        self.extra_file_name = '_temp'
        # for report contain one or multi testing items
        # default is single ( one verification)
        # change the name to _pXX for multi program
        # XX means the program selection number

        # 220907 add another variable call detail name for the
        # eff or I2C measurement
        # since save from each round of the eff test file
        self.detail_name = ''

        # the name can be change from main program(different user during program)
        self.flexible_name = ''

        # 221129 add the time stamp of the file name
        self.time_name = ''

        # assign one sheet is not raw out if efficiency test now used
        self.sh_temp = self.sh_volt_curr_cmd
        # the string indicate current items for file name or other adjustmenr reference
        # update the index when run verification start
        self.current_item_index = ''
        # identify if result book is close from the verification
        # open means can save file, close means skip end_of_file
        self.result_book_status = 'close'

        # 221209: add other file for extension
        self.last_report = self.sh_main.range('A5').value
        # check on the last_report with try except function

        # 220914 the instrument control variable
        self.inst_auto_selection = 1
        # this object is only used for self auto testing, no need to disable due to
        # conflict like eff_inst, so just set this to 1 directly,
        # reference this variable to B11 if going to mapped with auto control

        # result_book_trace change in the sub_program
        # update the result book trace
        self.full_result_name = self.new_file_name + \
            self.extra_file_name + self.detail_name
        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + self.detail_name + '.xlsx'

        # insturment parameter loading, to load the instrumenet paramenter
        # need to get the index of each item first
        # index need to change if adding new control parameter

        self.index_par_pre_con = 0
        self.index_GPIB_inst = 0
        self.index_general_other = 0
        self.index_pwr_inst = 0
        self.index_chroma_inst = 0
        self.index_src_inst = 0
        self.index_meter_inst = 0
        self.index_chamber_inst = 0
        self.index_IQ_scan = 0
        self.index_eff = 0
        self.index_general_test = 0

        self.index_par_pre_con = self.sh_main.range((3, 9)).value
        self.index_GPIB_inst = self.sh_main.range((4, 9)).value
        self.index_general_other = self.sh_main.range((5, 9)).value
        self.index_pwr_inst = self.sh_main.range((6, 9)).value
        self.index_chroma_inst = self.sh_main.range((3, 12)).value
        self.index_src_inst = self.sh_main.range((4, 12)).value
        self.index_meter_inst = self.sh_main.range((5, 12)).value
        self.index_chamber_inst = self.sh_main.range((6, 12)).value
        self.index_IQ_scan = self.sh_main.range((3, 15)).value
        self.index_eff = self.sh_main.range((4, 15)).value
        self.index_general_test = self.sh_main.range((5, 15)).value
        self.index_waveform_capture = self.sh_main.range((6, 15)).value

        # self.index_meter_inst = self.sh_main.range((5, 15)).value
        # self.index_chamber_inst = self.sh_main.range((6, 15)).value

        # index check put at the open result sheet

        # base on output format copied from the control book
        # start parameter initialization
        # pre- test condition settings
        self.pre_vin = self.sh_main.range(
            (self.index_par_pre_con + 1, 3)).value
        self.pre_vin_max = self.sh_main.range(
            (self.index_par_pre_con + 2, 3)).value
        self.pre_imax = self.sh_main.range(
            (self.index_par_pre_con + 3, 3)).value
        self.pre_test_en = self.sh_main.range(
            (self.index_par_pre_con + 4, 3)).value
        self.pre_sup_iout = self.sh_main.range(
            (self.index_par_pre_con + 5, 3)).value

        # load the GPIB address for the instrument
        # GPIB instrument list (address loading, name feed back)
        self.pwr_supply_addr = self.sh_main.range(
            (self.index_GPIB_inst + 1, 3)).value
        # met1v usually for voltage
        self.meter1_v_addr = self.sh_main.range(
            (self.index_GPIB_inst + 2, 3)).value
        # met2 usually for current
        self.meter2_i_addr = self.sh_main.range(
            (self.index_GPIB_inst + 3, 3)).value
        self.loader_addr = self.sh_main.range(
            (self.index_GPIB_inst + 4, 3)).value
        self.loader_src_addr = self.sh_main.range(
            (self.index_GPIB_inst + 5, 3)).value
        self.chamber_addr = self.sh_main.range(
            (self.index_GPIB_inst + 6, 3)).value
        self.scope_addr = self.sh_main.range(
            (self.index_GPIB_inst + 7, 3)).value
        self.pwr_bk_addr = self.sh_main.range(
            (self.index_GPIB_inst + 8, 3)).value
        # self.main_off_line = int(self.sh_main.range('A32').value)

        # initialization for all the object, based on the input parameter of the index

        # parameter setting for the power supply
        self.pwr_vset = self.sh_main.range((self.index_pwr_inst + 1, 3)).value
        self.pwr_iset = self.sh_main.range((self.index_pwr_inst + 2, 3)).value
        self.pwr_act_ch = self.sh_main.range(
            (self.index_pwr_inst + 3, 3)).value
        self.pwr_ini_state = self.sh_main.range(
            (self.index_pwr_inst + 4, 3)).value
        self.relay0_ch = self.sh_main.range((self.index_pwr_inst + 5, 3)).value
        self.relay6_ch = self.sh_main.range((self.index_pwr_inst + 6, 3)).value
        self.relay7_ch = self.sh_main.range((self.index_pwr_inst + 7, 3)).value
        # pre-increase for efficiency measurement
        self.pre_inc_vin = self.sh_main.range(
            (self.index_pwr_inst + 8, 3)).value
        # the setting for Vin calibration accuracy
        self.vin_diff_set = self.sh_main.range(
            (self.index_pwr_inst + 9, 3)).value

        # parameter setting for the chroma loader
        self.loader_act_ch = self.sh_main.range(
            (self.index_chroma_inst + 1, 3)).value
        self.loader_ini_mode = self.sh_main.range(
            (self.index_chroma_inst + 2, 3)).value
        self.loader_cal_offset_ELch = self.sh_main.range(
            (self.index_chroma_inst + 3, 3)).value
        self.loader_cal_offset_VCIch = self.sh_main.range(
            (self.index_chroma_inst + 4, 3)).value
        self.loader_ELch = self.sh_main.range(
            (self.index_chroma_inst + 5, 3)).value
        self.loader_ini_state = self.sh_main.range(
            (self.index_chroma_inst + 6, 3)).value
        self.loader_VCIch = self.sh_main.range(
            (self.index_chroma_inst + 7, 3)).value
        self.loader_cal_mode = self.sh_main.range(
            (self.index_chroma_inst + 8, 3)).value
        self.loader_cal_leakage_ELch = self.sh_main.range(
            (self.index_chroma_inst + 9, 3)).value
        self.loader_cal_leakage_VCIch = self.sh_main.range(
            (self.index_chroma_inst + 10, 3)).value

        # parameter setting for source meter
        self.src_vset = self.sh_main.range((self.index_src_inst + 1, 3)).value
        self.src_iset = self.sh_main.range((self.index_src_inst + 2, 3)).value
        self.src_ini_state = self.sh_main.range(
            (self.index_src_inst + 3, 3)).value
        self.src_ini_type = self.sh_main.range(
            (self.index_src_inst + 4, 3)).value
        self.src_clamp_ini = self.sh_main.range(
            (self.index_src_inst + 5, 3)).value

        # parameter setting for meter
        self.met_v_res = self.sh_main.range(
            (self.index_meter_inst + 1, 3)).value
        self.met_v_max = self.sh_main.range(
            (self.index_meter_inst + 2, 3)).value
        self.met_i_res = self.sh_main.range(
            (self.index_meter_inst + 3, 3)).value
        self.met_i_max = self.sh_main.range(
            (self.index_meter_inst + 4, 3)).value

        # parameter setting for chamber

        self.cham_tset_ini = self.sh_main.range(
            (self.index_chamber_inst + 1, 3)).value
        self.cham_ini_state = self.sh_main.range(
            (self.index_chamber_inst + 2, 3)).value
        self.cham_l_limt = self.sh_main.range(
            (self.index_chamber_inst + 3, 3)).value
        self.cham_h_limt = self.sh_main.range(
            (self.index_chamber_inst + 4, 3)).value
        self.cham_hyst = self.sh_main.range(
            (self.index_chamber_inst + 5, 3)).value

        # other control parameter
        # COM port parameter input
        self.mcu_com_addr = self.sh_main.range(
            self.index_general_other + 1, 3).value
        # general delay time
        self.wait_time = self.sh_main.range(
            self.index_general_other + 2, 3).value
        # the start point for the raw_out index
        self.raw_y_position_start = self.sh_main.range(
            self.index_general_other + 3, 3).value
        self.raw_x_position_start = self.sh_main.range(
            self.index_general_other + 4, 3).value
        self.book_off_finished = self.sh_main.range(
            self.index_general_other + 5, 3).value
        # plot pause control (1 is enable, 0 is disable)
        self.en_plot_waring = self.sh_main.range(
            self.index_general_other + 6, 3).value
        self.en_fully_auto = self.sh_main.range(
            self.index_general_other + 7, 3).value
        self.en_start_up_check = self.sh_main.range(
            self.index_general_other + 8, 3).value
        self.wait_small = self.sh_main.range(
            self.index_general_other + 9, 3).value

        # verification item: IQ parameter
        self.ISD_range = self.sh_main.range(
            self.index_IQ_scan + 1, 3).value
        self.sh_iq_scan_name = str(self.sh_main.range(
            self.index_IQ_scan + 2, 3).value)
        self.sh_sw_scan_name = str(self.sh_main.range(
            self.index_IQ_scan + 3, 3).value)

        # this is the sheet for IQ scan
        self.sh_iq_scan = self.wb.sheets(self.sh_iq_scan_name)
        # this is the sheet for wire scan
        self.sh_sw_scan = self.wb.sheets(self.sh_sw_scan_name)

        # verification item: eff control parameter
        self.channel_mode = self.sh_main.range(self.index_eff + 1, 3).value
        # SWIRE or I2C selected setting
        self.sw_i2c_select = self.sh_main.range(self.index_eff + 2, 3).value
        # if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
        # when control = 0, all channel used chroma's output mapping
        self.source_meter_channel = self.sh_main.range(
            self.index_eff + 3, 3).value
        # when control = 0, all channel used chroma's output mapping
        self.eff_chamber_en = self.sh_main.range(
            self.index_eff + 4, 3).value
        self.eff_single_file = self.sh_main.range(
            self.index_eff + 5, 3).value
        self.eff_rerun_en = self.sh_main.range(
            self.index_eff + 6, 3).value
        self.sh_volt_curr_cmd_name = str(self.sh_main.range(
            self.index_eff + 7, 3).value)
        self.sh_i2c_cmd_name = str(self.sh_main.range(
            self.index_eff + 8, 3).value)
        # 221129: change the sheet mapping of the command of efficiency

        # this is the sheet for efficiency testing command
        self.sh_volt_curr_cmd = self.wb.sheets(str(self.sh_volt_curr_cmd_name))
        # this is the sheet for I2C command
        self.sh_i2c_cmd = self.wb.sheets(str(self.sh_i2c_cmd_name))

        # verification item: general testing
        self.gen_chamber_en = self.sh_main.range(
            self.index_general_test + 1, 3).value
        self.gen_loader_en = self.sh_main.range(
            self.index_general_test + 2, 3).value
        self.gen_met_i_en = self.sh_main.range(
            self.index_general_test + 3, 3).value
        self.gen_volt_ch_amount = self.sh_main.range(
            self.index_general_test + 4, 3).value
        self.gen_pulse_i2x_en = self.sh_main.range(
            self.index_general_test + 5, 3).value
        self.gen_loader_ch_amount = self.sh_main.range(
            self.index_general_test + 6, 3).value
        self.gen_pwr_ch_amount = self.sh_main.range(
            self.index_general_test + 7, 3).value
        self.gen_pwr_i_set = self.sh_main.range(
            self.index_general_test + 8, 3).value
        self.gen_col_amount = self.sh_main.range(
            self.index_general_test + 9, 3).value
        self.single_test_mapped_general = self.sh_main.range(
            self.index_general_test + 10, 3).value

        # verification item: waveform capture object parameter in main
        self.pwr_select = int(self.sh_main.range(
            self.index_waveform_capture + 1, 3).value)
        self.scope_value = str(self.sh_main.range(
            self.index_waveform_capture + 2, 3).value)
        # change this setting to single verification control sheet
        # self.ripple_line_load = int(self.sh_main.range(
        #     self.index_waveform_capture + 1, 3).value)
        # add the loop control for each items
        self.single_test_mapped_wave = self.sh_main.range(
            self.index_waveform_capture + 3, 3).value

        # counteer is usually use c_ in opening

        # EFF_inst used
        self.c_avdd_load = self.sh_volt_curr_cmd.range('D1').value
        self.c_vin = self.sh_volt_curr_cmd.range('B1').value
        self.c_iload = self.sh_volt_curr_cmd.range('C1').value
        self.c_pulse = self.sh_volt_curr_cmd.range('E1').value
        self.c_i2c = self.sh_i2c_cmd.range('B1').value
        self.c_i2c_g = self.sh_i2c_cmd.range('D1').value
        self.c_avdd_single = self.sh_volt_curr_cmd.range('G1').value
        self.c_avdd_pulse = self.sh_volt_curr_cmd.range('H1').value
        self.c_tempature = self.sh_volt_curr_cmd.range('I1').value

        # IQ testing related
        self.c_iq = self.sh_iq_scan.range('C4').value
        self.iq_scaling = self.sh_iq_scan.range('C5').value

        # SWIRE_scan  related
        self.c_swire = self.sh_sw_scan.range('B2').value
        self.vin_set = self.sh_sw_scan.range('C6').value
        self.Iin_set = self.sh_sw_scan.range('E6').value
        self.EL_curr = self.sh_sw_scan.range('C7').value
        self.VCI_curr = self.sh_sw_scan.range('E7').value

        # efficiency test needed variable
        self.eff_done_sh = 1
        self.sub_sh_count = 0
        self.one_file_sheet_adj = 0

        # temp sheet name for the plot index
        # because there are different mode, no need specific channel name, just the positive and negative
        self.eff_temp = ''
        self.pos_temp = ''
        self.neg_temp = ''
        self.raw_temp = ''
        # 220825 add for vout and von regulation
        self.pos_pre_temp = ''
        self.neg_pre_temp = ''

        # the excel table gap for the data in raw sheet
        self.raw_gap = 4 + 10
        # setting of raw gap is the " gap + element "
        # single eff: Vin, Iin, Vout, Iout, Eff => 5 elements

        # active sheet (for the result and raw)
        self.sheet_active = ''
        # for efficiency
        self.raw_active = ''
        # for raw data
        self.vout_p_active = ''
        # for ELVDD, or AVDD
        self.vout_n_active = ''
        # for ELVSS
        self.vout_p_pre_active = ''
        # for VOP
        self.vout_n_pre_active = ''
        # for VON

        # 220926 format gen related control variable
        self.c_row_item = 0
        self.c_column_item = 0
        self.c_data_mea = 0
        self.c_ctrl_var1 = 0
        self.c_ctrl_var2 = 0
        self.c_ctrl_var4 = 0
        # fixed start point of the format gen (waveform element), (2, 5)
        self.format_start_x = 2
        self.format_start_y = 5
        # record the width and height from format gen and can be loaded to
        # 221102: the summary table at format gen control sheet, (13, 7) M7, L6 move
        self.summary_start_x = 13
        self.summary_start_y = 7
        # gap for the summary table from left to right
        self.summary_gap = 4
        self.ref_table_list = [None] * 100

        # waveform capture related testing
        self.wave_height = 93.8
        self.wave_width = 31.36
        # path need to be assign after every format gen finished
        self.wave_path = ''
        # the sheet name record for the saving waveform
        self.wave_sheet = ''
        # extra name for the testing conditon for waveform name
        self.wave_condition = ''
        # the full path used to load the waveform from HDD
        self.full_path = ''

        # =============
        # instrument control related

        # 20220502 instrument control parameter loading
        # use update sub program for continuous update the parameter from wb
        # position index (y-pos, x-pos)
        self.insref_pwr_y = self.sh_inst_ctrl.range('J5').value
        self.insref_load_y = self.sh_inst_ctrl.range('J6').value
        self.insref_met1_y = self.sh_inst_ctrl.range('J7').value
        self.insref_met2_y = self.sh_inst_ctrl.range('J8').value
        self.insref_src_y = self.sh_inst_ctrl.range('J9').value

        self.insref_pwr_x = self.sh_inst_ctrl.range('K5').value
        self.insref_load_x = self.sh_inst_ctrl.range('K6').value
        self.insref_met1_x = self.sh_inst_ctrl.range('K7').value
        self.insref_met2_x = self.sh_inst_ctrl.range('K8').value
        self.insref_src_x = self.sh_inst_ctrl.range('K9').value

        # self.insctl_refresh = self.sh_inst_ctrl.range('N4').value

        # variable to get the status update from main program

        self.pwr_connection_status = 0

        self.pwr_v_ch1_status = 0
        self.pwr_i_ch1_status = 0
        self.pwr_o_ch1_status = 0
        self.pwr_auto_ch1_status = 1

        self.pwr_v_ch2_status = 0
        self.pwr_i_ch2_status = 0
        self.pwr_o_ch2_status = 0
        self.pwr_auto_ch2_status = 1

        self.pwr_v_ch3_status = 0
        self.pwr_i_ch3_status = 0
        self.pwr_o_ch3_status = 0
        self.pwr_auto_ch3_status = 1

        self.load_connection_status = 0

        self.load_i_ch1_status = 0
        self.load_m_ch1_status = 0
        self.load_o_ch1_status = 0
        self.load_auto_ch1_status = 1

        self.load_i_ch2_status = 0
        self.load_m_ch2_status = 0
        self.load_o_ch2_status = 0
        self.load_auto_ch2_status = 1

        self.load_i_ch3_status = 0
        self.load_m_ch3_status = 0
        self.load_o_ch3_status = 0
        self.load_auto_ch3_status = 1

        self.load_i_ch4_status = 0
        self.load_m_ch4_status = 0
        self.load_o_ch4_status = 0
        self.load_auto_ch4_status = 1

        self.met1_connection_status = 0
        self.met1_mode_status = 0
        self.met1_level_status = 0
        self.met1_v_mea_status = 0
        self.met1_i_mea_status = 0

        self.met2_connection_status = 0
        self.met2_mode_status = 0
        self.met2_level_status = 0
        self.met2_v_mea_status = 0
        self.met2_i_mea_status = 0

        self.src_connection_status = 0
        self.src_mode_status = 0
        self.src_clamp_status = 0
        self.src_level_status = 0
        self.src_v_mea_status = 0
        self.src_i_mea_status = 0
        self.src_o_status = 0

        # self.inssts => instrument status related parameter (blue)
        # status output indexing (y-pos and x-pos)
        self.inssts_pwr_connection_y = int(self.insref_pwr_y)
        self.inssts_pwr_connection_x = int(self.insref_pwr_x) + 1

        self.inssts_pwr_refresh_y = int(self.insref_pwr_y) + 1
        self.inssts_pwr_refresh_x = int(self.insref_pwr_x) + 2

        self.inssts_pwr_serial_y = int(self.insref_pwr_y)
        self.inssts_pwr_serial_x = int(self.insref_pwr_x) + 5

        self.inssts_pwr_calibration_y = int(self.insref_pwr_y) + 1
        self.inssts_pwr_calibration_x = int(self.insref_pwr_x) + 5

        self.inssts_pwr_vset_ch1_y = int(self.insref_pwr_y) + 4
        self.inssts_pwr_vset_ch1_x = int(self.insref_pwr_x) + 1
        self.inssts_pwr_iset_ch1_y = int(self.insref_pwr_y) + 6
        self.inssts_pwr_iset_ch1_x = int(self.insref_pwr_x) + 1
        self.inssts_pwr_outs_ch1_y = int(self.insref_pwr_y) + 8
        self.inssts_pwr_outs_ch1_x = int(self.insref_pwr_x) + 1
        self.inssts_pwr_auto_ch1_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_auto_ch1_x = int(self.insref_pwr_x) + 1

        self.inssts_pwr_vset_ch2_y = int(self.insref_pwr_y) + 4
        self.inssts_pwr_vset_ch2_x = int(self.insref_pwr_x) + 3
        self.inssts_pwr_iset_ch2_y = int(self.insref_pwr_y) + 6
        self.inssts_pwr_iset_ch2_x = int(self.insref_pwr_x) + 3
        self.inssts_pwr_outs_ch2_y = int(self.insref_pwr_y) + 8
        self.inssts_pwr_outs_ch2_x = int(self.insref_pwr_x) + 3
        self.inssts_pwr_auto_ch2_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_auto_ch2_x = int(self.insref_pwr_x) + 3

        self.inssts_pwr_vset_ch3_y = int(self.insref_pwr_y) + 4
        self.inssts_pwr_vset_ch3_x = int(self.insref_pwr_x) + 5
        self.inssts_pwr_iset_ch3_y = int(self.insref_pwr_y) + 6
        self.inssts_pwr_iset_ch3_x = int(self.insref_pwr_x) + 5
        self.inssts_pwr_outs_ch3_y = int(self.insref_pwr_y) + 8
        self.inssts_pwr_outs_ch3_x = int(self.insref_pwr_x) + 5
        self.inssts_pwr_auto_ch3_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_auto_ch3_x = int(self.insref_pwr_x) + 5

        # position of index for channel status (for status color change)
        self.inssts_pwr_pos_ch1_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_pos_ch1_x = int(self.insref_pwr_x) + 0
        self.inssts_pwr_pos_ch2_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_pos_ch2_x = int(self.insref_pwr_x) + 2
        self.inssts_pwr_pos_ch3_y = int(self.insref_pwr_y) + 2
        self.inssts_pwr_pos_ch3_x = int(self.insref_pwr_x) + 4

        # DC loader status output index
        self.inssts_load_connection_y = int(self.insref_load_y)
        self.inssts_load_connection_x = int(self.insref_load_x) + 1

        self.inssts_load_refresh_y = int(self.insref_load_y) + 1
        self.inssts_load_refresh_x = int(self.insref_load_x) + 2

        self.inssts_load_iset_ch1_y = int(self.insref_load_y) + 4
        self.inssts_load_iset_ch1_x = int(self.insref_load_x) + 1
        self.inssts_load_mset_ch1_y = int(self.insref_load_y) + 6
        self.inssts_load_mset_ch1_x = int(self.insref_load_x) + 1
        self.inssts_load_outs_ch1_y = int(self.insref_load_y) + 8
        self.inssts_load_outs_ch1_x = int(self.insref_load_x) + 1
        self.inssts_load_auto_ch1_y = int(self.insref_load_y) + 2
        self.inssts_load_auto_ch1_x = int(self.insref_load_x) + 1

        self.inssts_load_iset_ch2_y = int(self.insref_load_y) + 4
        self.inssts_load_iset_ch2_x = int(self.insref_load_x) + 3
        self.inssts_load_mset_ch2_y = int(self.insref_load_y) + 6
        self.inssts_load_mset_ch2_x = int(self.insref_load_x) + 3
        self.inssts_load_outs_ch2_y = int(self.insref_load_y) + 8
        self.inssts_load_outs_ch2_x = int(self.insref_load_x) + 3
        self.inssts_load_auto_ch2_y = int(self.insref_load_y) + 2
        self.inssts_load_auto_ch2_x = int(self.insref_load_x) + 3

        self.inssts_load_iset_ch3_y = int(self.insref_load_y) + 4
        self.inssts_load_iset_ch3_x = int(self.insref_load_x) + 5
        self.inssts_load_mset_ch3_y = int(self.insref_load_y) + 6
        self.inssts_load_mset_ch3_x = int(self.insref_load_x) + 5
        self.inssts_load_outs_ch3_y = int(self.insref_load_y) + 8
        self.inssts_load_outs_ch3_x = int(self.insref_load_x) + 5
        self.inssts_load_auto_ch3_y = int(self.insref_load_y) + 2
        self.inssts_load_auto_ch3_x = int(self.insref_load_x) + 5

        self.inssts_load_iset_ch4_y = int(self.insref_load_y) + 4
        self.inssts_load_iset_ch4_x = int(self.insref_load_x) + 7
        self.inssts_load_mset_ch4_y = int(self.insref_load_y) + 6
        self.inssts_load_mset_ch4_x = int(self.insref_load_x) + 7
        self.inssts_load_outs_ch4_y = int(self.insref_load_y) + 8
        self.inssts_load_outs_ch4_x = int(self.insref_load_x) + 7
        self.inssts_load_auto_ch4_y = int(self.insref_load_y) + 2
        self.inssts_load_auto_ch4_x = int(self.insref_load_x) + 7

        # muti-meter status output index
        self.inssts_met1_connection_y = int(self.insref_met1_y)
        self.inssts_met1_connection_x = int(self.insref_met1_x) + 1
        self.inssts_met1_refresh_y = int(self.insref_met1_y) + 7
        self.inssts_met1_refresh_x = int(self.insref_met1_x) + 1
        self.inssts_met1_mset_y = int(self.insref_met1_y) + 3
        self.inssts_met1_mset_x = int(self.insref_met1_x) + 1
        self.inssts_met1_leve_y = int(self.insref_met1_y) + 5
        self.inssts_met1_leve_x = int(self.insref_met1_x) + 1
        self.inssts_met1_meav_y = int(self.insref_met1_y) + 7
        self.inssts_met1_meav_x = int(self.insref_met1_x) + 0
        self.inssts_met1_meai_y = int(self.insref_met1_y) + 9
        self.inssts_met1_meai_x = int(self.insref_met1_x) + 0

        self.inssts_met2_connection_y = int(self.insref_met2_y)
        self.inssts_met2_connection_x = int(self.insref_met2_x) + 1
        self.inssts_met2_refresh_y = int(self.insref_met2_y) + 7
        self.inssts_met2_refresh_x = int(self.insref_met2_x) + 1
        self.inssts_met2_mset_y = int(self.insref_met2_y) + 3
        self.inssts_met2_mset_x = int(self.insref_met2_x) + 1
        self.inssts_met2_leve_y = int(self.insref_met2_y) + 5
        self.inssts_met2_leve_x = int(self.insref_met2_x) + 1
        self.inssts_met2_meav_y = int(self.insref_met2_y) + 7
        self.inssts_met2_meav_x = int(self.insref_met2_x) + 0
        self.inssts_met2_meai_y = int(self.insref_met2_y) + 9
        self.inssts_met2_meai_x = int(self.insref_met2_x) + 0

        # source meter status output index

        self.inssts_src_connection_y = int(self.insref_src_y)
        self.inssts_src_connection_x = int(self.insref_src_x) + 1
        self.inssts_src_refresh_y = int(self.insref_src_y) + 9
        self.inssts_src_refresh_x = int(self.insref_src_x) + 1
        self.inssts_src_cset_y = int(self.insref_src_y) + 3
        self.inssts_src_cset_x = int(self.insref_src_x) + 1
        self.inssts_src_mset_y = int(self.insref_src_y) + 5
        self.inssts_src_mset_x = int(self.insref_src_x) + 1
        self.inssts_src_leve_y = int(self.insref_src_y) + 7
        self.inssts_src_leve_x = int(self.insref_src_x) + 1
        self.inssts_src_meav_y = int(self.insref_src_y) + 9
        self.inssts_src_meav_x = int(self.insref_src_x) + 0
        self.inssts_src_meai_y = int(self.insref_src_y) + 11
        self.inssts_src_meai_x = int(self.insref_src_x) + 0
        self.inssts_src_outs_y = int(self.insref_src_y) + 13
        self.inssts_src_outs_x = int(self.insref_src_x) + 1

        # finished the output index settings
        # since it will be fixed after table reference cell is set, only need to load once

        # self.insctl => instrument control related parameter (green)
        # input parameter loading

        # DC power supply table
        self.insctl_pwr_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 1, int(self.insref_pwr_x) + 1)).value
        self.insctl_pwr_serial = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y), int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_calibration = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 1, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_vset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_iset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_outs_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_vset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_iset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_outs_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_vset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_iset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_outs_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 4)).value

        # DC loader table
        self.insctl_load_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 1, int(self.insref_load_x) + 1)).value
        self.insctl_load_iset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 0)).value
        self.insctl_load_mset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 0)).value
        self.insctl_load_outs_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 0)).value
        self.insctl_load_iset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 2)).value
        self.insctl_load_mset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 2)).value
        self.insctl_load_outs_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 2)).value
        self.insctl_load_iset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 4)).value
        self.insctl_load_mset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 4)).value
        self.insctl_load_outs_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 4)).value
        self.insctl_load_iset_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 6)).value
        self.insctl_load_mset_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 6)).value
        self.insctl_load_outs_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 6)).value

        # multi-meter talbe

        self.insctl_met1_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 1, int(self.insref_met1_x) + 1)).value
        self.insctl_met1_mset = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 3, int(self.insref_met1_x) + 0)).value
        # the measurement level of the meter
        self.insctl_met1_leve = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 5, int(self.insref_met1_x) + 0)).value

        self.insctl_met2_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 1, int(self.insref_met2_x) + 1)).value
        self.insctl_met2_mset = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 3, int(self.insref_met2_x) + 0)).value
        # the measurement level of the meter
        self.insctl_met2_leve = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 5, int(self.insref_met2_x) + 0)).value

        # source-meter talbe

        self.insctl_src_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 1, int(self.insref_src_x) + 1)).value
        # setting of the clamp parameter
        self.insctl_src_cset = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 3, int(self.insref_src_x) + 0)).value
        self.insctl_src_mset = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 5, int(self.insref_src_x) + 0)).value
        # the setting level of the source meter
        self.insctl_src_leve = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 7, int(self.insref_src_x) + 0)).value
        # source meter ona and off control status
        self.insctl_src_outs = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 13, int(self.insref_src_x) + 0)).value

        # instrument control related
        # =============

        print('end of the parameter loaded')

        pass

    # sub_program of instrument check

    def para_update_pwr(self):
        # no need to input the parameter, check all the items based on the refrech time setting for each device
        # settings for table no need to use global, because re-load from excel every time,
        # but the status need to be global, since it will save the previous status
        # need to separate all the different instrument (there are different refresh rate)

        # call for parameter updatae
        # DC power supply table

        # global self.insctl_pwr_refresh
        # global self.insctl_pwr_serial
        # global self.insctl_pwr_calibration
        # global self.insctl_pwr_vset_ch1
        # global self.insctl_pwr_iset_ch1
        # global self.insctl_pwr_outs_ch1
        # global self.insctl_pwr_vset_ch2
        # global self.insctl_pwr_iset_ch2
        # global self.insctl_pwr_outs_ch2
        # global self.insctl_pwr_vset_ch3
        # global self.insctl_pwr_iset_ch3
        # global self.insctl_pwr_outs_ch3

        # DC power supply table
        self.insctl_pwr_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 1, int(self.insref_pwr_x) + 1)).value
        self.insctl_pwr_serial = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y), int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_calibration = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 1, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_vset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_iset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_outs_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 0)).value
        self.insctl_pwr_vset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_iset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_outs_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 2)).value
        self.insctl_pwr_vset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 4, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_iset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 6, int(self.insref_pwr_x) + 4)).value
        self.insctl_pwr_outs_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 8, int(self.insref_pwr_x) + 4)).value

        print('para_update_pwr')
        print('self.insctl_pwr_refresh = ' + str(self.insctl_pwr_refresh))
        print('self.insctl_pwr_serial = ' + str(self.insctl_pwr_serial))
        print('self.insctl_pwr_calibration = ' +
              str(self.insctl_pwr_calibration))
        print('self.insctl_pwr_vset_ch1 = ' + str(self.insctl_pwr_vset_ch1))
        print('self.insctl_pwr_iset_ch1 = ' + str(self.insctl_pwr_iset_ch1))
        print('self.insctl_pwr_outs_ch1 = ' + str(self.insctl_pwr_outs_ch1))
        print('self.insctl_pwr_vset_ch2 = ' + str(self.insctl_pwr_vset_ch2))
        print('self.insctl_pwr_iset_ch2 = ' + str(self.insctl_pwr_iset_ch2))
        print('self.insctl_pwr_outs_ch2 = ' + str(self.insctl_pwr_outs_ch2))
        print('self.insctl_pwr_vset_ch3 = ' + str(self.insctl_pwr_vset_ch3))
        print('self.insctl_pwr_iset_ch3 = ' + str(self.insctl_pwr_iset_ch3))
        print('self.insctl_pwr_outs_ch3 = ' + str(self.insctl_pwr_outs_ch3))

        pass

    def status_update_pwr(self):
        # update the status to the related excel table
        self.sh_inst_ctrl.range((self.inssts_pwr_connection_y,
                                 self.inssts_pwr_connection_x)).value = self.pwr_connection_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_refresh_y, self.inssts_pwr_refresh_x)).value = self.insctl_pwr_refresh
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_serial_y, self.inssts_pwr_serial_x)).value = self.insctl_pwr_serial
        self.sh_inst_ctrl.range((self.inssts_pwr_calibration_y,
                                 self.inssts_pwr_calibration_x)).value = self.insctl_pwr_calibration
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_vset_ch1_y, self.inssts_pwr_vset_ch1_x)).value = self.pwr_v_ch1_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_iset_ch1_y, self.inssts_pwr_iset_ch1_x)).value = self.pwr_i_ch1_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_outs_ch1_y, self.inssts_pwr_outs_ch1_x)).value = self.pwr_o_ch1_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_auto_ch1_y, self.inssts_pwr_auto_ch1_x)).value = self.pwr_auto_ch1_status

        self.sh_inst_ctrl.range(
            (self.inssts_pwr_vset_ch2_y, self.inssts_pwr_vset_ch2_x)).value = self.pwr_v_ch2_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_iset_ch2_y, self.inssts_pwr_iset_ch2_x)).value = self.pwr_i_ch2_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_outs_ch2_y, self.inssts_pwr_outs_ch2_x)).value = self.pwr_o_ch2_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_auto_ch2_y, self.inssts_pwr_auto_ch2_x)).value = self.pwr_auto_ch2_status

        self.sh_inst_ctrl.range(
            (self.inssts_pwr_vset_ch3_y, self.inssts_pwr_vset_ch3_x)).value = self.pwr_v_ch3_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_iset_ch3_y, self.inssts_pwr_iset_ch3_x)).value = self.pwr_i_ch3_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_outs_ch3_y, self.inssts_pwr_outs_ch3_x)).value = self.pwr_o_ch3_status
        self.sh_inst_ctrl.range(
            (self.inssts_pwr_auto_ch3_y, self.inssts_pwr_auto_ch3_x)).value = self.pwr_auto_ch3_status

        pass

    def para_update_load(self):

        # global self.insctl_load_refresh
        # global self.insctl_load_iset_ch1
        # global self.insctl_load_mset_ch1
        # global self.insctl_load_outs_ch1
        # global self.insctl_load_iset_ch2
        # global self.insctl_load_mset_ch2
        # global self.insctl_load_outs_ch2
        # global self.insctl_load_iset_ch3
        # global self.insctl_load_mset_ch3
        # global self.insctl_load_outs_ch3
        # global self.insctl_load_iset_ch4
        # global self.insctl_load_mset_ch4
        # global self.insctl_load_outs_ch4

        # DC loader table
        self.insctl_load_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 1, int(self.insref_load_x) + 1)).value
        self.insctl_load_iset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 0)).value
        self.insctl_load_mset_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 0)).value
        self.insctl_load_outs_ch1 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 0)).value
        self.insctl_load_iset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 2)).value
        self.insctl_load_mset_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 2)).value
        self.insctl_load_outs_ch2 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 2)).value
        self.insctl_load_iset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 4)).value
        self.insctl_load_mset_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 4)).value
        self.insctl_load_outs_ch3 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 4)).value
        self.insctl_load_iset_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 4, int(self.insref_load_x) + 6)).value
        self.insctl_load_mset_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 6, int(self.insref_load_x) + 6)).value
        self.insctl_load_outs_ch4 = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 8, int(self.insref_load_x) + 6)).value

        print('para_update_load')
        print('self.insctl_load_refresh = ' + str(self.insctl_load_refresh))
        print('self.insctl_load_iset_ch1 = ' + str(self.insctl_load_iset_ch1))
        print('self.insctl_load_mset_ch1 = ' + str(self.insctl_load_mset_ch1))
        print('self.insctl_load_outs_ch1 = ' + str(self.insctl_load_outs_ch1))
        print('self.insctl_load_iset_ch2 = ' + str(self.insctl_load_iset_ch2))
        print('self.insctl_load_mset_ch2 = ' + str(self.insctl_load_mset_ch2))
        print('self.insctl_load_outs_ch2 = ' + str(self.insctl_load_outs_ch2))
        print('self.insctl_load_iset_ch3 = ' + str(self.insctl_load_iset_ch3))
        print('self.insctl_load_mset_ch3 = ' + str(self.insctl_load_mset_ch3))
        print('self.insctl_load_outs_ch3 = ' + str(self.insctl_load_outs_ch3))
        print('self.insctl_load_iset_ch4 = ' + str(self.insctl_load_iset_ch4))
        print('self.insctl_load_mset_ch4 = ' + str(self.insctl_load_mset_ch4))
        print('self.insctl_load_outs_ch4 = ' + str(self.insctl_load_outs_ch4))

        pass

    def status_update_load(self):

        # update the status to the related excel table
        self.sh_inst_ctrl.range((self.inssts_load_connection_y,
                                 self.inssts_load_connection_x)).value = self.load_connection_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_refresh_y, self.inssts_load_refresh_x)).value = self.insctl_load_refresh

        # ch1
        self.sh_inst_ctrl.range(
            (self.inssts_load_iset_ch1_y, self.inssts_load_iset_ch1_x)).value = self.load_i_ch1_status
        self.sh_inst_ctrl.range((self.inssts_load_mset_ch1_y,
                                 self.inssts_load_mset_ch1_x)).value = self.load_m_ch1_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_outs_ch1_y, self.inssts_load_outs_ch1_x)).value = self.load_o_ch1_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_auto_ch1_y, self.inssts_load_auto_ch1_x)).value = self.load_auto_ch1_status

        # ch2
        self.sh_inst_ctrl.range(
            (self.inssts_load_iset_ch2_y, self.inssts_load_iset_ch2_x)).value = self.load_i_ch2_status
        self.sh_inst_ctrl.range((self.inssts_load_mset_ch2_y,
                                 self.inssts_load_mset_ch2_x)).value = self.load_m_ch2_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_outs_ch2_y, self.inssts_load_outs_ch2_x)).value = self.load_o_ch2_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_auto_ch2_y, self.inssts_load_auto_ch2_x)).value = self.load_auto_ch2_status

        # ch3
        self.sh_inst_ctrl.range(
            (self.inssts_load_iset_ch3_y, self.inssts_load_iset_ch3_x)).value = self.load_i_ch3_status
        self.sh_inst_ctrl.range((self.inssts_load_mset_ch3_y,
                                 self.inssts_load_mset_ch3_x)).value = self.load_m_ch3_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_outs_ch3_y, self.inssts_load_outs_ch3_x)).value = self.load_o_ch3_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_auto_ch3_y, self.inssts_load_auto_ch3_x)).value = self.load_auto_ch3_status

        # ch4
        self.sh_inst_ctrl.range(
            (self.inssts_load_iset_ch4_y, self.inssts_load_iset_ch4_x)).value = self.load_i_ch4_status
        self.sh_inst_ctrl.range((self.inssts_load_mset_ch4_y,
                                 self.inssts_load_mset_ch4_x)).value = self.load_m_ch4_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_outs_ch4_y, self.inssts_load_outs_ch4_x)).value = self.load_o_ch4_status
        self.sh_inst_ctrl.range(
            (self.inssts_load_auto_ch4_y, self.inssts_load_auto_ch4_x)).value = self.load_auto_ch4_status

        pass

    def para_update_met1(self):

        # global self.insctl_met1_refresh
        # global self.insctl_met1_mset
        # global self.insctl_met1_leve

        self.insctl_met1_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 1, int(self.insref_met1_x) + 1)).value
        self.insctl_met1_mset = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 3, int(self.insref_met1_x) + 0)).value
        # the measurement level of the meter
        self.insctl_met1_leve = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 5, int(self.insref_met1_x) + 0)).value

        print('para_update_met1')
        print('self.insctl_met1_refresh = ' + str(self.insctl_met1_refresh))
        print('self.insctl_met1_mset = ' + str(self.insctl_met1_mset))
        print('self.insctl_met1_leve = ' + str(self.insctl_met1_leve))

        pass

    def status_update_met1(self):

        # global met1_mode_status
        # global met1_level_status
        # global met1_v_mea_status
        # global met1_i_mea_status

        # if self.insctl_met1_mset == 0:
        #     # enter the voltage measurement mode

        #     met1_i_mea_status = 'NA'
        #     met1_mode_status = 'votlage'
        #     met1_level_status = self.insctl_met1_leve
        #     pass

        # elif self.insctl_met1_mset == 1:
        #     # enter the current measurement mode

        #     met1_v_mea_status = 'NA'
        #     met1_mode_status = 'current'
        #     met1_level_status = self.insctl_met1_leve
        #     pass

        # 220521 some of the status variable comes from the main program, so need other variable save for the result
        self.sh_inst_ctrl.range(
            (self.inssts_met1_connection_y, self.inssts_met1_connection_x)).value = self.met1_connection_status
        self.sh_inst_ctrl.range((self.inssts_met1_refresh_y,
                                 self.inssts_met1_refresh_x)).value = self.insctl_met1_refresh
        self.sh_inst_ctrl.range(
            (self.inssts_met1_mset_y, self.inssts_met1_mset_x)).value = self.met1_mode_status
        self.sh_inst_ctrl.range(
            (self.inssts_met1_leve_y, self.inssts_met1_leve_x)).value = self.met1_level_status
        self.sh_inst_ctrl.range(
            (self.inssts_met1_meav_y, self.inssts_met1_meav_x)).value = self.met1_v_mea_status
        self.sh_inst_ctrl.range(
            (self.inssts_met1_meai_y, self.inssts_met1_meai_x)).value = self.met1_i_mea_status

        pass

    def para_update_met2(self):

        # global self.insctl_met2_refresh
        # global self.insctl_met2_mset
        # global self.insctl_met2_leve

        self.insctl_met2_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 1, int(self.insref_met2_x) + 1)).value
        self.insctl_met2_mset = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 3, int(self.insref_met2_x) + 0)).value
        # the measurement level of the meter
        self.insctl_met2_leve = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 5, int(self.insref_met2_x) + 0)).value

        print('para_update_met2')
        print('self.insctl_met2_refresh = ' + str(self.insctl_met2_refresh))
        print('self.insctl_met2_mset = ' + str(self.insctl_met2_mset))
        print('self.insctl_met2_leve = ' + str(self.insctl_met2_leve))

        pass

    def status_update_met2(self):

        # global met2_mode_status
        # global met2_level_status
        # global met2_v_mea_status
        # global met2_i_mea_status

        # if self.insctl_met2_mset == 0:
        #     # enter the voltage measurement mode

        #     met2_i_mea_status = 'NA'
        #     met2_mode_status = 'votlage'
        #     met2_level_status = self.insctl_met2_leve
        #     pass

        # elif self.insctl_met2_mset == 1:
        #     # enter the current measurement mode

        #     met2_v_mea_status = 'NA'
        #     met2_mode_status = 'current'
        #     met2_level_status = self.insctl_met2_leve
        #     pass

        self.sh_inst_ctrl.range(
            (self.inssts_met2_connection_y, self.inssts_met2_connection_x)).value = self.met2_connection_status
        self.sh_inst_ctrl.range((self.inssts_met2_refresh_y,
                                 self.inssts_met2_refresh_x)).value = self.insctl_met2_refresh
        self.sh_inst_ctrl.range(
            (self.inssts_met2_mset_y, self.inssts_met2_mset_x)).value = self.met2_mode_status
        self.sh_inst_ctrl.range(
            (self.inssts_met2_leve_y, self.inssts_met2_leve_x)).value = self.met2_level_status
        self.sh_inst_ctrl.range(
            (self.inssts_met2_meav_y, self.inssts_met2_meav_x)).value = self.met2_v_mea_status
        self.sh_inst_ctrl.range(
            (self.inssts_met2_meai_y, self.inssts_met2_meai_x)).value = self.met2_i_mea_status

        pass

    def para_update_src(self):

        # global self.insctl_src_refresh
        # global self.insctl_src_cset
        # global self.insctl_src_mset
        # global self.insctl_src_leve
        # global self.insctl_src_outs

        self.insctl_src_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 1, int(self.insref_src_x) + 1)).value
        # setting of the clamp parameter
        self.insctl_src_cset = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 3, int(self.insref_src_x) + 0)).value
        self.insctl_src_mset = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 5, int(self.insref_src_x) + 0)).value
        # the setting level of the source meter
        self.insctl_src_leve = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 7, int(self.insref_src_x) + 0)).value
        # source meter ona and off control status
        self.insctl_src_outs = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 13, int(self.insref_src_x) + 0)).value

        print('para_update_src')
        print('self.insctl_src_refresh = ' + str(self.insctl_src_refresh))
        print('self.insctl_src_cset = ' + str(self.insctl_src_cset))
        print('self.insctl_src_mset = ' + str(self.insctl_src_mset))
        print('self.insctl_src_leve = ' + str(self.insctl_src_leve))
        print('self.insctl_src_outs = ' + str(self.insctl_src_outs))

        pass

    def status_update_src(self):

        self.sh_inst_ctrl.range(
            (self.inssts_src_connection_y, self.inssts_src_connection_x)).value = self.src_connection_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_refresh_y, self.inssts_src_refresh_x)).value = self.insctl_src_refresh
        self.sh_inst_ctrl.range(
            (self.inssts_src_cset_y, self.inssts_src_cset_x)).value = self.src_clamp_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_mset_y, self.inssts_src_mset_x)).value = self.src_mode_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_leve_y, self.inssts_src_leve_x)).value = self.src_level_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_meav_y, self.inssts_src_meav_x)).value = self.src_v_mea_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_meai_y, self.inssts_src_meai_x)).value = self.src_i_mea_status
        self.sh_inst_ctrl.range(
            (self.inssts_src_outs_y, self.inssts_src_outs_x)).value = self.src_o_status

        pass

    def check_refresh(self):
        # this sub is used prevent the dead loop of latch refresh setting

        # global self.insctl_pwr_refresh
        # global self.insctl_load_refresh
        # global self.insctl_met1_refresh
        # global self.insctl_met2_refresh
        # global self.insctl_src_refresh

        self.insctl_pwr_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_pwr_y) + 1, int(self.insref_pwr_x) + 1)).value
        self.insctl_load_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_load_y) + 1, int(self.insref_load_x) + 1)).value
        self.insctl_met1_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met1_y) + 1, int(self.insref_met1_x) + 1)).value
        self.insctl_met2_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_met2_y) + 1, int(self.insref_met2_x) + 1)).value
        self.insctl_src_refresh = self.sh_inst_ctrl.range(
            (int(self.insref_src_y) + 1, int(self.insref_src_x) + 1)).value

        # also need to update the refresh status to the excel table
        # so people know if refresh command is enter or not

        self.sh_inst_ctrl.range(
            (self.inssts_pwr_refresh_y, self.inssts_pwr_refresh_x)).value = self.insctl_pwr_refresh

        self.sh_inst_ctrl.range(
            (self.inssts_load_refresh_y, self.inssts_load_refresh_x)).value = self.insctl_load_refresh

        self.sh_inst_ctrl.range((self.inssts_met1_refresh_y,
                                 self.inssts_met1_refresh_x)).value = self.insctl_met1_refresh

        self.sh_inst_ctrl.range((self.inssts_met2_refresh_y,
                                 self.inssts_met2_refresh_x)).value = self.insctl_met2_refresh

        self.sh_inst_ctrl.range(
            (self.inssts_src_refresh_y, self.inssts_src_refresh_x)).value = self.insctl_src_refresh

        print('check_refresh')
        print('self.insctl_pwr_refresh = ' + str(self.insctl_pwr_refresh))
        print('self.insctl_load_refresh = ' + str(self.insctl_load_refresh))
        print('self.insctl_met1_refresh = ' + str(self.insctl_met1_refresh))
        print('self.insctl_met2_refresh = ' + str(self.insctl_met2_refresh))
        print('self.insctl_src_refresh = ' + str(self.insctl_src_refresh))

        pass

    def open_result_book(self, keep_last=0):

        # before open thr result book to check index, first check and correct
        # index (will be update to the obj_main)
        if self.result_book_status == 'close':
            self.index_check()
            if keep_last == 0:
                # define new result workbook
                self.wb_res = xw.Book()
            else:
                try:
                    self.wb_res = xw.Book(
                        f'c:\\py_gary\\test_excel\\{self.last_report}.xlsx')
                except:
                    print(
                        'error mapping with last report, check name and will process new book\n g is cute')
                    # define new result workbook
                    self.wb_res = xw.Book()
                    # change the setting back to normal if there are no proper last report to open
                    keep_last = 0

            # create reference sheet (for sheet position)
            # sh_ref is the index for result sheet
            # sh_ref_condition is for testing condition and setting
            # all the reference sheet will delete after the program finished
            self.sh_ref = self.wb_res.sheets.add('ref_sh')
            # self.sh_ref_condition = self.wb_res.sheets.add('ref_sh2')
            # delete the extra sheet from new workbook, difference from version
            if keep_last == 0:
                self.wb_res.sheets('工作表1').delete()
            else:
                # delete the original main sheet and replace with the new one
                self.wb_res.sheets('main').delete()

            # copy the main sheets to new book
            self.sh_main.copy(self.sh_ref)
            # assign sheet to the new sheets in result book
            self.sh_main = self.wb_res.sheets('main')

            # 221113: assign the setting file name to the related file
            self.sh_main.range('B14').value = self.wb.name

            # 221129: add the start time stamp to record time needed
            self.sh_main.range('B5').value = self.time_stamp()

            # for the other sheet rather than main, will decide to copy to result
            # or not depends on verification item is used or not
            self.extra_file_name = '_temp'
            self.result_book_status = 'open'
            self.excel_save()
            pass
        else:
            print('result book already open, may have errors')
            time.sleep(3)
            pass

        pass

    def end_of_file(self, multi_items):
        # at the end of test, delete the reference sheet and save the file
        # 220907 change name from end_of_test to end_of file, since the operation
        # to cut file may be needed during single test

        # 220914 if result book already close, skip all the action
        if self.result_book_status == 'open':

            self.sh_ref.delete()
            # if self.eff_single_file == 0:
            #     self.sh_temp.delete()

            # self.sh_ref_condition.delete()
            # update the result book trace
            # extra file name should be update by the last item or the single item
            if multi_items == 1:
                # using multi item extra file name
                self.extra_file_name = '_p' + \
                    str(int(self.program_group_index))

            if self.eff_single_file == 1 and self.current_item_index == 'eff':
                print(self.detail_name)
                # input()
                self.detail_name = '_eff all in 1'

            # 221129: add the time stamp of the file
            self.time_name = self.time_stamp()
            self.sh_main.range('B6').value = self.time_name

            self.result_book_trace = self.excel_temp + \
                self.new_file_name + self.extra_file_name + \
                self.detail_name + self.flexible_name + '_' + self.time_name + '.xlsx'
            self.full_result_name = self.new_file_name + \
                self.extra_file_name + self.detail_name + \
                self.flexible_name + '_' + self.time_name
            self.wb_res.save(self.result_book_trace)

            # 221209: add the copy to summary book after finished the testing?

            if self.book_off_finished == 1:
                self.wb_res.close()
                pass

            # to reset the sheet after file finished and turn off
            self.sheet_reset()
            self.detail_name = ''
            self.flexible_name = ''
            self.time_name = ''
            self.extra_file_name = '_temp'
            self.new_file_name = str(self.sh_main.range('B8').value)
            self.full_result_name = self.new_file_name + \
                self.extra_file_name + self.detail_name + self.flexible_name + self.time_name
            self.result_book_trace = self.excel_temp + \
                self.new_file_name + self.extra_file_name + \
                self.detail_name + self.flexible_name + self.time_name + '.xlsx'

            # reset the sheet count of the one file efficiency when end of file
            self.one_file_sheet_adj = 0

            self.result_book_status = 'close'
            pass
        else:
            print('there should be error, no result book to save')
            print("Love for g is trying to help her when she needs")
            time.sleep(3)
            pass

        pass

    def flexible_naming(self, name_string):
        '''
        is able to change the file name at the end during program process\n
        but only reserve for Grace(root) access XD
        '''

        # this can be the flexible name of the file name call by main object
        # to have different file name without changing the excel
        self.flexible_name = str(name_string)

        pass

    def time_stamp(self):
        '''
        to add the time stamp the end of file, need
        '''

        now = datetime.now()  # current date and time

        year = now.strftime("%Y")
        print("year:", year)

        month = now.strftime("%m")
        print("month:", month)

        day = now.strftime("%d")
        print("day:", day)

        time = now.strftime("%H:%M:%S")
        print("time:", time)

        date_time = now.strftime("%Y-%m-%d, %H:%M:%S")
        print("date and time:", date_time)

        date_stamp = now.strftime("%Y_%m_%d_%H_%M")
        print('the stamp: ' + date_stamp)

        return date_stamp

    def excel_save(self):
        # only save, not change the result book trace
        # should be only the temp file during program operation
        self.wb_res.save(self.result_book_trace)

        # also update the program exit control for checking
        self.check_program_exit()

        pass

    def inst_name_sheet(self, nick_name, full_name):
        # definition of sub program may not need the self, but definition of class will need the self
        # self is usually used for internal parameter of class
        # this function will get the nick name and full name from main and update to the sheet
        # based on the nick name

        # 220902 operate after main is change, and result will be correct

        if nick_name == 'PWR1':
            self.sh_main.range(
                (self.index_GPIB_inst + 1, 4)).value = full_name

        elif nick_name == 'MET1':
            self.sh_main.range(
                (self.index_GPIB_inst + 2, 4)).value = full_name

        elif nick_name == 'MET2':
            self.sh_main.range(
                (self.index_GPIB_inst + 3, 4)).value = full_name

        elif nick_name == 'LOAD1':
            self.sh_main.range(
                (self.index_GPIB_inst + 4, 4)).value = full_name

        elif nick_name == 'LOADSR':
            self.sh_main.range(
                (self.index_GPIB_inst + 5, 4)).value = full_name

        elif nick_name == 'chamber':
            self.sh_main.range(
                (self.index_GPIB_inst + 6, 4)).value = full_name

        elif nick_name == 'scope':
            self.sh_main.range(
                (self.index_GPIB_inst + 7, 4)).value = full_name

        elif nick_name == 'bkpwr':
            self.sh_main.range(
                (self.index_GPIB_inst + 8, 4)).value = full_name

        pass

    def sheet_reset(self):
        # this sheet reset the all the sheet variable assignment to the original sheet
        # in main_obj, which used for the re-run program
        # just copoy from the sheet assignment

        # after choosing the workbook, define the main sheet to load parameter
        self.sh_main = self.wb.sheets('main')

        # only the instrument control will be still mapped to the original excel
        # since inst_ctrl is no needed to copy to the result sheet
        self.sh_inst_ctrl = self.wb.sheets('inst_ctrl')

        # other way to define sheet:
        # this is the format for the efficiency result
        ex_sheet_name = 'raw_out'
        self.sh_raw_out = self.wb.sheets(ex_sheet_name)
        # this is the sheet for efficiency testing command
        self.sh_volt_curr_cmd = self.wb.sheets(self.sh_volt_curr_cmd_name)
        # this is the sheet for I2C command
        self.sh_i2c_cmd = self.wb.sheets(self.sh_i2c_cmd_name)
        # this is the sheet for IQ scan
        self.sh_iq_scan = self.wb.sheets(self.sh_iq_scan_name)
        # this is the sheet for wire scan
        self.sh_sw_scan = self.wb.sheets(self.sh_sw_scan_name)

        pass

    def index_check(self):
        # this sub program used to check the index setting for the excel input
        # prevent logic error of the wrong indexing of program parameter
        check_str = 'settings'
        index_correction = 0
        item_array = np.full([11], 1)
        check_ctrl = self.sum_array(item_array)

        while (check_ctrl > 0):
            # while there are items not pass, need to re-run check parameter
            # porcess

            if self.sh_main.range((int(self.index_par_pre_con), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_par_pre_con check done')
                item_array[0] = 0
                print(item_array)
                pass
            else:
                print('index_par_pre_con check fail')
                print('correct value in: (3, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 9)).value = index_correction
                self.index_par_pre_con = index_correction
                pass

            if self.sh_main.range((int(self.index_GPIB_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_GPIB_inst check done')
                item_array[1] = 0
                print(item_array)
                pass
            else:
                print('index_GPIB_inst check fail')
                print('correct value in: (4, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 9)).value = index_correction
                self.index_GPIB_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_general_other), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_general_other check done')
                item_array[2] = 0
                print(item_array)
                pass
            else:
                print('index_general_other check fail')
                print('correct value in: (5, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 9)).value = index_correction
                self.index_general_other = index_correction
                pass

            if self.sh_main.range((int(self.index_pwr_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_pwr_inst check done')
                item_array[3] = 0
                print(item_array)
                pass
            else:
                print('index_pwr_inst check fail')
                print('correct value in: (6, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((6, 9)).value = index_correction
                self.index_pwr_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_chroma_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_chroma_inst check done')
                item_array[4] = 0
                print(item_array)
                pass
            else:
                print('index_chroma_inst check fail')
                print('correct value in: (3, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 12)).value = index_correction
                self.index_chroma_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_src_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_src_inst check done')
                item_array[5] = 0
                print(item_array)
                pass
            else:
                print('index_src_inst check fail')
                print('correct value in: (4, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 12)).value = index_correction
                self.index_src_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_meter_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_meter_inst check done')
                item_array[6] = 0
                print(item_array)
                pass
            else:
                print('index_meter_inst check fail')
                print('correct value in: (5, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 12)).value = index_correction
                self.index_meter_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_chamber_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_chamber_inst check done')
                item_array[7] = 0
                print(item_array)
                pass
            else:
                print('index_chamber_inst check fail')
                print('correct value in: (6, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((6, 12)).value = index_correction
                self.index_chamber_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_IQ_scan), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_IQ_scan check done')
                item_array[8] = 0
                print(item_array)
                pass
            else:
                print('index_IQ_scan check fail')
                print('correct value in: (3, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 15)).value = index_correction
                self.index_IQ_scan = index_correction
                pass

            if self.sh_main.range((int(self.index_eff), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_eff check done')
                item_array[9] = 0
                print(item_array)
                pass
            else:
                print('index_eff check fail')
                print('correct value in: (4, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 15)).value = index_correction
                self.index_eff = index_correction
                pass
            # update while loop value
            check_ctrl = self.sum_array(item_array)

            if self.sh_main.range((int(self.index_general_test), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('general_test check done')
                item_array[10] = 0
                print(item_array)
                pass
            else:
                print('general_test check fail')
                print('correct value in: (5, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 15)).value = index_correction
                self.index_eff = index_correction
                pass
            # update while loop value
            check_ctrl = self.sum_array(item_array)

            if self.sh_main.range((int(self.index_waveform_capture), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('waveform_capture check done')
                item_array[10] = 0
                print(item_array)
                pass
            else:
                print('waveform_capture check fail')
                print('correct value in: (6, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 15)).value = index_correction
                self.index_eff = index_correction
                pass
            # update while loop value
            check_ctrl = self.sum_array(item_array)

            pass

        print('the index correction finished!')
        # save the wb for the new index correction settings
        # the wb in excel object is mapped to the obj_main file
        self.wb.save()
        pass

    def sum_array(self, arr):
        # initialize a variable
        # to store the sum
        # while iterating through
        # the array later
        sum = 0

        # iterate through the array
        # and add each element to the sum variable
        # one at a time
        for i in arr:
            sum = sum + i

        return (sum)
        pass

    def sim_mode_delay(self, wait_time, wait_small):
        # reduce the delay time for the simulation mode
        self.wait_time = wait_time
        self.wait_time = wait_small
        pass

    # SWIRE request sub-program

    def ideal_v_table(self, c_swire):
        ideal_v_res = self.sh_sw_scan.range((11 + c_swire, 3)).value
        return ideal_v_res

    #  efficiency testing sub-program

    def build_file(self, detail_name):

        # 220907 mapped variable with excel object
        self.detail_name = detail_name
        wb = self.wb
        wb_res = self.wb_res
        result_sheet_name = 'raw_out'
        new_file_name = self.new_file_name
        excel_trace = self.excel_temp
        channel_mode = self.channel_mode
        c_avdd_load = self.c_avdd_load
        sh_ref = self.sh_ref
        self.wb.sheets('raw_out').copy(self.sh_ref)
        self.sh_raw_out = wb_res.sheets('raw_out')
        # this sheet mapped to raw out at eff_inst
        sh_org_tab2 = self.sh_raw_out
        # this sheet mapped to volage and current command
        sh_org_tab = self.sh_volt_curr_cmd
        sh_ref = self.sh_ref
        sheet_arry = self.sheet_arry

        # cpoy the sheet to the result book, will be set at the eff_obj
        # 220907: here is only for the book generation, must call the sheet
        # generation before call the build file
        # assign for the result sheet to the excel object will also be done in the
        # sheet gen of eff_obj

        # update the file name, but not update file name untill the file is finished

        # save the result book and turn off the control book
        wb_res.save(self.result_book_trace)
        # wb.close()
        # close control books

        # base on output format copied from the control book
        # start parameter initialization

        # used array to setup the result sheet
        # res_sheet_array = np.zeros(100)
        # this method not working, skip this time

        # here is to generate the sheet
        # counter selection for sheet generation: EL power only one cycle needed (IAVDD= 0)
        if channel_mode == 0 or channel_mode == 1:
            # both only EL or only AVDD just one time
            c_sheet_copy = 1
            if channel_mode == 1:
                sub_sh_count = 4

                # eff + raw + AVDD regulation + Vout regulation
                # 220825 add vout sheet
            elif channel_mode == 0:
                sub_sh_count = 6
                # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation
        elif channel_mode == 2:
            c_sheet_copy = c_avdd_load
            sub_sh_count = 6
            # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation

        self.sub_sh_count = sub_sh_count
        # 220911 update the sub sh count after confirm the parameter

        x_sheet_copy = 0
        sh_temp = sh_org_tab2
        self.sh_temp = sh_temp
        # this loop build the extra sheet needed in the program
        # there are one raw data sheet and fixed format summarize table
        # need to build both efficiency and load regulation summarize table in this loop
        # 3-channel efficiency
        while x_sheet_copy < c_sheet_copy:
            # issue: if x_sheet_copy == 0:

            if channel_mode == 2 or channel_mode == 0:
                # sheet needed for 3 chand only EL are the same
                # just define the different sheet name

                # load AVDD current parameter
                excel_temp = str(sh_org_tab.range(3 + x_sheet_copy, 4).value)

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'EFF_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'EFF'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                if channel_mode == 2:
                    # add another sheet for the raw data of each AVDD current
                    sheet_temp = 'RAW_I_AVDD=' + excel_temp + 'A'
                else:
                    sheet_temp = 'RAW'
                # raw data sheet no need example format, can use empty sheet
                sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
                wb_res.sheets.add(sheet_temp)
                # =======
                # 220825 explanation added: since the sheet of raw data doesn't have specific
                # format and input needed, add the sheet directly, no need to copy
                # this is the reason why it's different with other sheet generation
                # to add the Vout and Von load regulation, use the format in excel raw_out
                # and it's general format for the regulation and plot function in VBA

                # =======
                # add another sheet for the ELVDD data of each AVDD current
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'ELVDD_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'ELVDD'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheset name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
                sh_org_tab2.name = sheet_temp
                # =======

                # =======
                # add another sheet for the ELVSS data of each AVDD current
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'ELVSS_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'ELVSS'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
                sh_org_tab2.name = sheet_temp
                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'Vop_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'Vop'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 4] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'Von_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'Von'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 5] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

            elif channel_mode == 1:
                # sheet build up for only AVDD

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'EFF'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                # add another sheet for the raw data of each AVDD current
                sheet_temp = 'RAW'
                # raw data sheet no need example format, can use empty sheet
                sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
                wb_res.sheets.add(sheet_temp)
                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'AVDD'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'Vout'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======
                self.sub_sh_count = sub_sh_count

            x_sheet_copy = x_sheet_copy + 1

        # assign the sheet to the raw out in resuslt sheet
        self.sh_temp = wb_res.sheets('raw_out')
        sh_temp.delete()
        # don't need the original raw output format, remove the output

    def message_box(self, content_str, title_str, auto_exception=0, box_type=0):
        '''
        message box function
        auto_exception is for waveform capture, will bypass fully auto setting in global setting \n
        boxtype(mpaaed with return value): 0-only confirm\n
        1-confirm: 1, cancel: 2
        2-stop: 3, re-try: 4, skip: 5
        3-yes: 6, no: 7, cancel: 2
        4-yes: 6, no: 7
        '''
        content_str = str(content_str)
        title_str = str(title_str)
        msg_res = 7
        # won't skip if not enter the result update
        if self.en_fully_auto == 0 or auto_exception == 1:
            msg_res = win32api.MessageBox(0, content_str, title_str, box_type)
        # 0 to 3 is different type of message box and can sen different return value
        # detail check on the internet
        print('msg box call~~ ')
        print('P.S Grace is cute! ~ ')

        return msg_res

    def eff_rerun(self):
        # this program check the status of the excel file eff_re-run block
        # and update the eff_done to restart efficienct testing
        # from the main, this sub will run if eff_done is already 1

        # 220913 new re-run setting for eff
        # check if the re-efficiency is needed for verification
        # jump the window if needed to re-run, and turn off if not
        # check the rerun_en first and load the verification re-run
        # after confirm the jump wondow
        # only run re-run process when re-run EN is 1
        if self.eff_rerun_en == 1:

            msg_res = win32api.MessageBox(
                0, 're-run is set to 1, update re-run_en to deecide to sotp in next round or not', 'Updating setting of re-run')

            eff_reset_temp = self.sh_main.range('B13').value
            while eff_reset_temp == 0:

                self.program_exit = self.sh_main.range('B12').value
                eff_reset_temp = self.sh_main.range('B13').value
                print('wait for re-run, update command and setup then set re-run to 1')
                print('the program will start again')
                print('break the loop by a kiss from Grace')
                time.sleep(2)
                print('just kidding~ . Set B12 in main to 0 to break')
                time.sleep(10)

                if self.program_exit == 0:
                    break
                pass

            eff_reset_temp = self.sh_main.range('B13').value

            if eff_reset_temp == 1:
                self.eff_done_sh = 0
                # reset to 0 if eff sheet is ready to re-run
                # also need to set te input blank back to 0
                self.sh_main.range('B13').value = 0
                self.sh_main.range('C13').value = 0
                # other wise there will be infinite loop

                # update the re-run control variable
                self.eff_rerun_en = self.sh_main.range(
                    self.index_eff + 6, 3).value

                # also need to re-assign the mapping sheet to Eff_inst
                # the sheet assignment is gone after finished one round
                self.sheet_reset()

                pass
            else:
                # no need for the action of changing the reset status
                pass

        return self.eff_done_sh

    def program_status(self, status_string):
        # transfer to the string for following operation
        # if you need to modify in sub-program, need to use global definition
        # global vin_status
        # global i_avdd_status
        # global i_el_status
        # global sw_i2c_status
        status_sting_sub = str(status_string)
        self.sh_main.range((3, 2)).value = status_sting_sub
        # Vin status
        self.sh_main.range('F3').value = self.vin_status
        # I_AVDD status
        self.sh_main.range('F4').value = self.i_avdd_status
        # I_EL staatus
        self.sh_main.range('F5').value = self.i_el_status
        # SW_I2C status
        self.sh_main.range('F6').value = self.sw_i2c_status

        print('status_update: ' + status_sting_sub)
        print(str(self.vin_status) + '-' + str(self.i_avdd_status) +
              '-' + str(self.i_el_status) + '-' + str(self.sw_i2c_status))
        # use for debugging for the program status update
        # input()
        pass

    def act_sheet_loaded(self):
        #  to consider to put the update of sheet assign to here or stay in eff_obj
        pass

    def eff_calculated(self):
        self.value_eff = ((self.value_elvdd - self.value_elvss) * self.value_iel +
                          self.value_avdd * self.value_iavdd) / (self.value_vin * self.value_iin)
        return self.value_eff

    def sheet_adj_for_eff(self, x_avdd):
        # used to adjust the sheet name to prevent conflict when
        # building in single file

        # re-name the sheet with new name

        x_sub_sh_count = 0
        while x_sub_sh_count < self.sub_sh_count:
            index = self.sub_sh_count * x_avdd + x_sub_sh_count
            target_sheet = self.wb_res.sheets(self.sheet_arry[index])
            new_sheet_name = str(self.one_file_sheet_adj) + \
                '_' + self.sheet_arry[index]
            target_sheet.name = new_sheet_name
            self.sheet_arry[index] = new_sheet_name

            # also need to update operating condition to each sheet
            self.condition_note = self.extra_file_name + self.detail_name
            if x_sub_sh_count == 0:
                target_sheet.range('M2').value = 'operating condition'
            target_sheet.range('M3').value = self.condition_note

            x_sub_sh_count = x_sub_sh_count + 1

        self.one_file_sheet_adj = self.one_file_sheet_adj + 1
        pass

    # used for the loading the data to related excel sheet and blank

    def data_latch(self, data_name, mea_res, x_vin, x_iload, value_i_offset1, value_i_offset2):
        raw_gap = self.raw_gap
        channel_mode = self.channel_mode
        # define the globa variable for eff calculation
        # global value_elvdd
        # global value_elvss
        # global value_avdd
        # global value_iin
        # global value_vin
        # global value_iel
        # global value_iavdd

        # global bypass_measurement_flag
        # first to check if the bypass flag raise ~
        # set measurement result to 0 if the bpass flag is enable
        # if bypass_measurement_flag == 1:
        #     mea_res = '0'

        if data_name == 'vin':
            # vin only record in the raw data
            self.raw_active.range((11 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_vin = float(mea_res)

        elif data_name == 'iin':
            # iin only record in the raw data
            self.raw_active.range((12 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_iin = float(mea_res)

        elif data_name == 'elvdd':
            # elvdd record in the raw data, elvdd regulation
            self.raw_active.range((13 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            # sheet_active.range((25 + x_iload, 3 + x_vin)).value = lo.atof(mea_res)
            if channel_mode == 0 or channel_mode == 2:
                self.vout_p_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)
            self.value_elvdd = float(mea_res)

        elif data_name == 'elvss':
            # elvss record in the raw data, elvss regulation
            self.raw_active.range((14 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            if channel_mode == 0 or channel_mode == 2:
                self.vout_n_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)
            self.value_elvss = float(mea_res)

        elif data_name == 'i_el':
            # i_el only record in the raw data
            self.raw_active.range((15 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res) - value_i_offset1
            self.value_iel = float(mea_res) - value_i_offset1

        elif data_name == 'avdd':
            # avvdd record in the raw data, avvdd regulation
            self.raw_active.range((16 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_avdd = float(mea_res)
            # the raw data of AVDD need to record no matter in AVDD only mode or the 3-ch mode
            # the selection of EL only or not is decidde inthe main program
            # here is only for the choice of AVDD regulation
            if channel_mode == 1:
                # only need to record the regulation when is operating for AVDD only mode
                self.vout_p_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)

        elif data_name == 'i_avdd':
            # i_avdd only record in the raw data
            self.raw_active.range((17 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res) - value_i_offset2
            self.value_iavdd = float(mea_res) - value_i_offset2

        elif data_name == 'eff':
            # eff record in the raw data
            self.raw_active.range((18 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.sheet_active.range((25 + x_iload, 3 + x_vin)
                                    ).value = lo.atof(mea_res)

        elif data_name == 'vop':
            # vop record in the raw data
            self.raw_active.range((19 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.vout_p_pre_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)

        elif data_name == 'von':
            # von record in the raw data
            self.raw_active.range((20 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            if channel_mode == 0 or channel_mode == 2:
                self.vout_n_pre_active.range((25 + x_iload, 3 + x_vin)
                                             ).value = lo.atof(mea_res)

        # clear the bypass flag every time enter data latch function
        # bypass_measurement_flag = 0

    # the instrument update check for excel
    def check_inst_update(self):

        pass

    # to plot the result in different sheet
    def plot_single_sheet(self, v_cnt, i_cnt, sheet_n):

        book_n = str(self.full_result_name) + '.xlsx'
        # plot the sheet based on the input sheet name and element length
        excel.Application.Run("obj_main.xlsm!gary_chart",
                              v_cnt, i_cnt, sheet_n, book_n, self.raw_y_position_start, self.raw_x_position_start)
        print('the plot of ' + str(sheet_n) +
              ' in book ' + str(book_n) + ' is finished')
        pass

    def check_program_exit(self):
        # the sub add for checking the program exit
        #  can be uased to skip the loop and prevent dead loop
        self.progrm_exit = self.sh_main.range('B12').value
        if self.progrm_exit == 0:
            self.turn_inst_off = 1
            # get out the program after saving the temp file
            print('always stop when Grace want to talk, forever and ever')
            time.sleep(0.20)
            if self.ready_to_off == 1:
                # this will be set to 1 after the object finished the reset of MCU and turn off instrument
                sys.exit()

        pass

    # sub program for the format generation

    def table_return(self):
        # need to recover this sheet: self.excel_ini.sh_ref_table
        self.sh_ref_table = self.wb.sheets('table')

        pass

    # sub program for waveform capture

    def scope_capture(self, target_sheet, range_index, default_trace=0, left=0, top=0, width=0, height=0):
        '''
        capture the waveform from the excel \n
        key in left to 0 to keep the original dimension and no need to input other
        '''

        if default_trace == 0.5:
            # this selection is reserve for the test mode
            default_trace = 'c:\\py_gary\\test_excel\\test_pic.png'

        if default_trace == 0:

            default_trace = self.full_path

        # if the parameter not going to use, need to define the default value, dimension can be felxible input

        # here is to call the VBA function and get the capture from the scope
        # will need to reference the scope library and see if it can be cover from
        # python only, should be easier

        # left is used to decide if need to adjust dimentions of cell
        if left == 0:
            # no need to adjust the dimension
            target_sheet.pictures.add(str(default_trace), left=range_index.left,
                                      top=range_index.top, width=range_index.width, height=range_index.height)

            pass
        else:

            target_sheet.pictures.add(str(default_trace), float(left),
                                      float(top), float(width), float(height))

            pass

        pass

    def get_nth_key(self, dictionary, n=0):

        if n < 0:
            n += len(dictionary)
        for i, key in enumerate(dictionary.keys()):
            if i == n:
                return key
        raise IndexError("dictionary index out of range")

        pass

    def get_nth_value(self, dictionary, n=0):

        if n < 0:
            n += len(dictionary)
        for i, key in enumerate(dictionary.values()):
            if i == n:
                return key
        raise IndexError("dictionary index out of range")

        pass

    def wave_info_update(self, **kwargs):
        '''
        this function should be call by 'run_verification' and input the information of testing in dictionary type:\n
        EX: Vin= v_target, I_load = i_load ......
        '''

        # generate the wveform naming string
        k1 = len(kwargs)
        x = 0
        self.wave_condition = '_'
        while x < k1:
            a = list(kwargs)[x]
            b = list(kwargs.values())[x]

            self.wave_condition = self.wave_condition + \
                str(a) + '-' + str(b) + '_'

            x = x + 1
            pass

        pass

    def float_gene(self, input, scaling=1000, digit=2):
        '''
        transfer the digit of float \n
        input can be string or float \n
        default scaling to mV
        '''
        a = float(input)
        a = a * scaling
        if digit == 0:
            b = float("{:.0f}".format(a))
        elif digit == 1:
            b = float("{:.1f}".format(a))
        elif digit == 2:
            b = float("{:.2f}".format(a))
        elif digit == 3:
            b = float("{:.3f}".format(a))
        elif digit == 4:
            b = float("{:.4f}".format(a))
        print(b)
        return b

    def sum_table_gen(self, ind_x, ind_y, x_axis=0, y_axis=0, content=0, sheet=0):
        '''
        this function is used to generate the summary table of testing result \n
        ind_x, ind_y are the index coordinate of table; status: 'build' or 'fill';
        content: string input; x_axis , y_axis are the coordinate for build input;
        count is the fill input; sheet is default set to mapped fomat gen sheet,
        change with verification items
        '''

        # if sheet set ot 0 is to used original control sheet
        if sheet == 0:
            sheet = self.sh_format_gen
            print('default sheet selected for g')

        sheet.range((ind_y + y_axis, ind_x + x_axis)).value = content

        pass

    def eff_para_re_load(self, new_VI_com=0, new_i2C_ctrl=0):
        '''
        the function is used to improve the control sheet extention of efficiency measurement\n
        send the name of new sheet into the function before run verification, and it can replace the control parameter before start
        easire to switch the command without changing the table

        '''

        # EFF_inst used
        self.c_avdd_load = self.sh_volt_curr_cmd.range('D1').value
        self.c_vin = self.sh_volt_curr_cmd.range('B1').value
        self.c_iload = self.sh_volt_curr_cmd.range('C1').value
        self.c_pulse = self.sh_volt_curr_cmd.range('E1').value
        self.c_i2c = self.sh_i2c_cmd.range('B1').value
        self.c_i2c_g = self.sh_i2c_cmd.range('D1').value
        self.c_avdd_single = self.sh_volt_curr_cmd.range('G1').value
        self.c_avdd_pulse = self.sh_volt_curr_cmd.range('H1').value
        self.c_tempature = self.sh_volt_curr_cmd.range('I1').value

        pass


if __name__ == '__main__':
    #  the testing code for this file object

    import datetime

    test_mode = 2.5

    excel = excel_parameter('obj_main')
    if test_mode == 0:
        input()

        excel2 = excel_parameter('other_testing_condition')

        input()

        # 221020
        # testing for the input of picture from file to excel
        # this is the second step for scope capture
        # need to open the test excel in folder

        pass

    elif test_mode == 1:

        wb_test = xw.Book('c:\\py_gary\\test_excel\\test.xlsx')

        test_sh = wb_test.sheets('test_sh')
        image_range = test_sh.range('C10')

        test_sh.pictures.add('c:\\py_gary\\test_excel\\test_pic.png', left=image_range.left,
                             top=image_range.top, width=image_range.width, height=image_range.height)

        print('end of add picture')

        pass

    elif test_mode == 1.5:
        # testing for the input parameter of the sub program
        def test(a1=0, *args, **kwargs):

            print(a1)

            print(args)
            k0 = len(args)
            if k0 > 0:
                print('this is args 1' + str(args[0]))
                print('this is args 2' + str(args[1]))

            print(kwargs)
            k1 = len(kwargs)
            if k1 > 0:
                print('this is kwargs 1 ' + str(kwargs['para1']))
                print('this is kwargs 2 ' + str(kwargs['para2']))

            print('end')

            pass

        def get_time_stamp():

            time = datetime.datetime.now()
            print(time)

            return time

        a = get_time_stamp()
        print(a)

        test('test for many parameter - 1', 'ab',
             'cd', para1=3, para2='par2')

        test('test for many parameter - 2')

        pass

    elif test_mode == 2:

        wb_test = xw.Book('c:\\py_gary\\test_excel\\test.xlsx')
        test_sh = wb_test.sheets('test_sh')
        image_range = test_sh.range('C10')

        # these can't be change from, only loaded from the cell object
        a = image_range.height
        b = image_range.width
        c = image_range.top
        d = image_range.left

        excel.scope_capture(
            'c:\\py_gary\\test_excel\\test_pic.png', test_sh, image_range)

        pass

    elif test_mode == 2.5:
        # this example discuss about how ot use dictionary index

        colors = {"blue": "5", "red": "6", "yellow": "8"}

        first_key = list(colors)[0]
        first_val = list(colors.values())[0]

        def get_nth_key(dictionary, n=0):

            if n < 0:
                n += len(dictionary)
            for i, key in enumerate(dictionary.keys()):
                if i == n:
                    return key
            raise IndexError("dictionary index out of range")

            pass

        def get_nth_value(dictionary, n=0):

            if n < 0:
                n += len(dictionary)
            for i, key in enumerate(dictionary.values()):
                if i == n:
                    return key
            raise IndexError("dictionary index out of range")

            pass

        return_key = get_nth_key(colors, 2)
        print(return_key)

        return_value = get_nth_value(colors, 2)
        print(return_value)

        pass
