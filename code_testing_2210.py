# this file is mainly for the testing of coding


import locale as lo


class test_calass():

    def __init__(self):

        pass

    def scope_ch(self):

        self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
        self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

        for i in range(1, 8+1):
            if i == 1:
                temp_dict = self.ch_c1

            a = temp_dict['ch_view']
            print(f'app.Acquisition.C{i}.View = {a}')
            print(f'app.Acquisition.C{i}.View = {temp_dict["ch_view"]}')

        pass

    def two_dim_dict(self):

        self.p1 = {"param": "max", "source": "C1", "view": "TRUE"}
        self.p2 = {"param": "min", "source": "C2", "view": "FALSE"}

        self.mea_ch = {"P1": self.p1, "P2": self.p2, }

        # multi layer of the list (dictionary)

        temp = list(self.mea_ch)[0]
        temp3 = list(self.mea_ch.values())[0]
        trmp2 = list(list(self.mea_ch.values())[0])[0]
        temp4 = (list(self.mea_ch.values())[0])["param"]

        print(
            f"app.Measure.{list(self.mea_ch)[0]}.ParamEngine = '{(list(self.mea_ch.values())[0])['param']}'")

        pass

    def float_format(self):

        # this fuction used to test format output for the float
        a = 0.123456789
        dig = 5
        b = float("{:.2f}".format(a))
        # c = float(f"{:.{dig}f}".format(a)) => error format
        d = "{:.4f}".format(a)
        e = lo.atof("{:.2f}".format(a))

        print(b)

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


t_s = test_calass()
testing_index = 1

if testing_index == 0:
    print('a')
    # from 0-9 => < 10 and start from 0 (like array)
    for i in range(10):
        print(i, end=' ')
        # not change line

    for i in range(1, 1+8):
        print(i)

        t_s.scope_ch()

elif testing_index == 1:
    # testing for the 2 dimension dictionary

    t_s.two_dim_dict()
    t_s.float_format()
    t_s.float_gene('0.123456789', 100)
    t_s.float_gene('0.123456789', digit=0)
    t_s.float_gene('0.123456789', digit=1)
    t_s.float_gene('0.123456789', digit=2)
    t_s.float_gene('0.123456789', digit=3)
    t_s.float_gene('0.123456789', digit=4)

    pass

elif testing_index == 1:

    pass
