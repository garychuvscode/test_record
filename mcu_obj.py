# build from 220829
# this file is to setup an object contain all the general used parameter or the general used
# sub program, and this object can be inheritance by the other verification object
# this may help with the generation of general used data and help for
# the decouple of entire system

import time
import pyvisa
rm = pyvisa.ResourceManager()

#  not only one class for all the stuff, separate for few class definition


class MCU_control ():

    def __init__(self, sim_mcu0, com_addr0):
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0
        if self.com_addr == 100:
            # set MCU to simulation mode if the addr is set to 100
            self.sim_mcu = 0

        # because MCU will be separate with GPIB for implementation and test
        self.mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
        # meter channel indicator: 0: Vin, 1: AVDD, 2: OVDD, 3: OVSS, 4: VOP, 5: VON
        # array mpaaing for the relay control
        self.meter_ch_ctrl = 0
        self.array_rst = '00'

        self.pulse1 = 0
        self.pulse2 = 0
        self.mode_set = 4
        # mode sequence: 1-4: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1) => default normal

        # i2c register and data (single byte data); slave address is fixed in the MCU
        # here only support for the register and data change
        self.reg_i2c = ''
        self.data_i2c = ''

        #  UART command string
        self.uart_cmd_str = ''

        # different mode used in the operation
        self.mcu_mode_swire = 1
        self.mcu_mode_sw_en = 3
        self.mcu_mode_I2C = 4
        self.mcu_mode_8_bit_IO = 5
        self.mcu_mode_pat_gen_py = 6
        self.mcu_mode_pat_gen_encode = 7
        self.mcu_mode_pat_gen_direct = 8

        # MCU mapping for different mode control in 2553
        # MCU mapping for RA_GPIO control is mode 5
        # both mode 1 and 2 should defined as dual SWIRE, need to send two pulse command at one time
        # need to be build in the control sheet
        self.wait_time = 0.5
        self.wait_small = 0.2
        # waiting time duration, can be adjust with MCU C code

        pass

    def mcu_write(self, index):
        # the MCU write used to generate the command string and send the command out
        # but the related control parameter need to define in the main program
        # here is only used to reduce the code of generate the string
        # and the MCU UART sending command

        if index == 'swire':
            self.uart_cmd_str = (chr(self.mcu_mode_swire) +
                                 chr(int(self.pulse1)) + chr(int(self.pulse2)))
            # for the SWIRE mode of 2553, there are 2 pulse send to the MCU and DUT
            # pulse amount is from 1 to 255, not sure if 0 will have error or not yet
            # 20220121
            print(
                f'2 swire pulse command in MCU {self.pulse1} and {self.pulse2}')
        elif index == 'en_sw':
            self.uart_cmd_str = (chr(self.mcu_mode_sw_en) +
                                 chr(int(self.mode_set)) + chr(1))
            # for the EN SWIRE control mode, need to handle the recover to normal mode (EN, SW) = (1, 1)
            # at the end of application
            # this mode only care about the first data ( 0-4 )
            print(f'mode command in MCU, mode{self.mode_set}')
        elif index == 'relay':
            self.uart_cmd_str = (
                chr(self.mcu_mode_8_bit_IO) + self.mcu_cmd_arry[self.meter_ch_ctrl])
            # assign relay to related channel after function called
            # channel index is from golbal variable
            print(
                f'relay command in MCU ch {self.mcu_cmd_arry[self.meter_ch_ctrl]} from {self.meter_ch_ctrl}')

        elif index == 'i2c':
            self.uart_cmd_str = (chr(self.mcu_mode_I2C) +
                                 str(self.reg_i2c) + str(self.data_i2c))
            print(
                f'i2C command in MCU reg: {self.reg_i2c} data: {self.data_i2c}')
            # send mapped i2c command out from MCU
        elif index == 'grace':
            # reserve for the special function of MCU write
            # reset the relay or the GPIO channel
            self.uart_cmd_str = (
                chr(self.mcu_mode_8_bit_IO) + '00')
            print('call grace in MCU, reset the I/O channel to all low')
            print('grace is so cute~')
        else:
            print('wrong command on MCU, double check g')
            print('re-send the PMIC mode for instead')
            self.uart_cmd_str = (chr(self.mcu_mode_sw_en) +
                                 chr(int(self.mode_set)) + chr(1))

            #  if sending the wrong string, gaive the message through terminal

        # print the command going to send before write to MCU, used for debug
        print('the string is at next line')
        print(self.uart_cmd_str)
        if self.sim_mcu == 1:
            # now is real mode, output the MCU command from COM port
            # self.mcu_com.write(self.uart_cmd_str)
            # 221113: since for the old version board may have error,
            # use query instead (MCU will also reply the command)
            return_str = self.mcu_com.query(self.uart_cmd_str)
            print(f'command sent with return {return_str}')
        else:
            print('now is sending the MCU command with below string:')
            print(str(index))
            print(self.uart_cmd_str)

        # give some response time for the UART command send and MCU action
        time.sleep(self.wait_small)

    def com_open(self):
        # this function is used to open the com port of the MCU
        # this will be set independentlly in each object
        print('now the COM port is on')
        uart_cmd_str = "COM" + str(int(self.com_addr))
        print(uart_cmd_str)
        if self.sim_mcu == 1:
            self.mcu_com = rm.open_resource(uart_cmd_str)
        else:
            print('open COM port but bypass the real operation')

        pass

    def com_close(self):
        # after the verification is finished, reset all the condition
        # to initial and turn off the communication port
        self.back_to_initial()
        print('the MCU will turn off')
        if self.sim_mcu == 1:
            self.mcu_com.close()
        else:
            print('the com port is turn off now')

        pass

    def back_to_initial(self):
        # this sub program used to set all the MCU condition to initial
        # to change the initial setting, just modify the items from here

        # MCU will be in normal mode (EN, SW) = (1, 1) => 4
        self.pmic_mode(4)
        # the relay channel also reset to the default

        print('command accept to reset the MCU')
        pass

    def pulse_out(self, pulse_1, pulse_2):
        '''
        pulse need to be less than 255
        '''
        # pulse should be within 255 (8bit data limitation)
        if pulse_1 > 255:
            pulse_1 = 255
        if pulse_2 > 255:
            pulse_2 = 255
        # pulse = 0 is the issue wait for workaround
        # need to be change in MCU code
        # if pulse_1 == 0:
        #     pulse_1 = 255
        # if pulse_2 == 0:
        #     pulse_2 = 255
        self.pulse1 = pulse_1
        self.pulse2 = pulse_2
        if (self.pulse1 == 0) and (self.pulse2 == 0):
            print('pulse 0 0 is send, MCU no action')
            print('cute Grace!')
            pass
        elif self.pulse1 == 0:
            if self.pulse2 != 0:
                self.pulse1 = self.pulse2
                self.mcu_write('swire')
                pass
            pass
        elif self.pulse2 == 0:
            self.pulse2 = self.pulse1
            self.mcu_write('swire')
            pass
        elif (self.pulse1 != 0) and (self.pulse2 != 0):
            self.mcu_write('swire')
            pass

        pass

    def pmic_mode(self, mode_index):
        '''
        (EN,SW) or (EN2, EN1) \n
        1:(0,0); 2:(0,1); 3:(1,0); 4:(1,1)
        '''
        # mode index should be in 1-4
        if mode_index < 1 or mode_index > 4:
            mode_index = 1
            # turn off if error occur
        self.mode_set = mode_index
        self.mcu_write('en_sw')
        pass

    def relay_ctrl(self, channel_index):
        '''
        index array: ['01', '02', '04', '08', '10', '20', '40', '80']\n
        from 0 to 7 \n
        MCU IO 2.0 - 2.7
        '''
        self.meter_ch_ctrl = int(channel_index)
        self.mcu_write('relay')
        pass

    def i2c_single_write(self, register_index, data_index):
        self.reg_i2c = register_index
        self.data_i2c = data_index
        self.mcu_write('i2c')

    # to update the implementation of other function
    # think about what is needed from the MCU operation
    # update later

    pass

#  testing items for the MCU control object
# 220901 other MCU function wait for added


if __name__ == '__main__':

    mcu_cmd_arry_I2C_org = ['C3', '1C', 'C5', '05', 'C6', '1B', 'C4', '2B']
    mcu_cmd_arry_I2C_0p1 = ['C3', '1D', 'C5', '06', 'C6', '1C', 'C4', '2C']
    select_i2c = 0
    # select I2C => 0 is org, 1 is 0p1 array
    mcu = MCU_control(1, 3)
    mcu.com_open()

    mcu.pulse_out(10, 25)
    input()
    mcu.pulse_out(0, 10)
    input()
    mcu.pulse_out(5, 0)
    input()
    mcu.pulse_out(0, 0)
    input()

    mcu.pmic_mode(1)
    input()
    mcu.pmic_mode(4)
    input()
    mcu.pmic_mode(2)
    input()
    mcu.pmic_mode(3)
    input()
    mcu.back_to_initial()
    input()
    x_relay = 0
    while x_relay < 8:
        # relay channel sweep for the test mode
        mcu.relay_ctrl(x_relay)
        print(str(mcu.mcu_cmd_arry[x_relay]))
        time.sleep(0.1)
        x_relay = x_relay + 1

    input()
    x_i2c = 0
    while x_i2c < 4:
        if select_i2c == 1:
            # uart_cmd_str = (chr(int(
            #     mcu.mcu_mode_I2C)) + mcu_cmd_arry_I2C_0p1[2 * x_i2c] + mcu_cmd_arry_I2C_0p1[2 * x_i2c + 1])
            mcu.i2c_single_write(
                mcu_cmd_arry_I2C_0p1[2 * x_i2c], mcu_cmd_arry_I2C_0p1[2 * x_i2c + 1])
            print('register: ' + str(mcu_cmd_arry_I2C_0p1[2 * x_i2c]) + ', date: ' + str(
                mcu_cmd_arry_I2C_0p1[2 * x_i2c + 1]))
        else:
            # choose different array for the I2C command output
            # uart_cmd_str = (chr(int(
            #     mcu.mcu_mode_I2C)) + mcu_cmd_arry_I2C_org[2 * x_i2c] + mcu_cmd_arry_I2C_org[2 * x_i2c + 1])
            mcu.i2c_single_write(
                mcu_cmd_arry_I2C_org[2 * x_i2c], mcu_cmd_arry_I2C_org[2 * x_i2c + 1])
            print('register: ' + str(mcu_cmd_arry_I2C_org[2 * x_i2c]) + ', date: ' + str(
                mcu_cmd_arry_I2C_org[2 * x_i2c + 1]))
        # print(uart_cmd_str + str(select_i2c))

        time.sleep(mcu.wait_small)
        print('mcu I2C finished')
        input()

        x_i2c = x_i2c + 1
    input()
    mcu.back_to_initial()
    print('mcu ready to turn off')
    input()
    mcu.com_close()

    pass
