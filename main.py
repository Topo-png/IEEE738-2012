# -*- coding: utf-8 -*-
"""
Created on Sun Apr  3 22:33:02 2022

@author: shuck
"""

"""
        import conductor data frame
        import configuration data frame
        convert configuration and store in new data frame to units specified by 
        covert conductor and store in new data frame to units specified by config dataframe
            conductor OD 
            conductor resistance distance
            conductor resistance temperature
            conductor normal temperature rating
            conductor emergency temperature rating

        create df_temperature_ambient_range
            row 0-n - based on ambient temp increments, low, and upper range, set in configuration
            col 0 - original temperature values
            co1 1 - original temperature units
            col 2 - converted temperature values
            col 3 - converted temperature units C

        create df_temperature_conductor_range
            row 0-n - temperature range based on # of steps, self.conductor_temp_steps
            col 0 - original temperature values
            co1 1 - original temperature units
            col 2 - converted temperature values
            col 3 - converted temperature units C

        create df_normal_output_values
            row 0-n - results based on each calculation required
            col 0 - normal current rating
            col 1 - qc used in calculation
            col 2 - qc0
            col 3 - qc1
            col 4 - qc2
            col 5 - qr 
            col 6 - qs (calculated only once)
            col 7 - conductor resistance
            col 8 - uf
            col 9 - pf
            col 10 - kf
            col 11 - etc... 
        """

import numpy as np
import pandas as pd
import datetime
import time

import UnitConversion as unc
uc = unc.UnitConvert()

ver = 'v0.1.0'
debug = False

dir_config = 'Sample/config-sample.xlsx'  # location of configuration file
dir_conductor = 'Sample/Conductor_Prop-Sample.xlsx'  # location of conductor file

degree_sign = u'\N{DEGREE SIGN}'

class IEEE738:

    def __init__(self):
        self.true_to_standard = True
        self.metric_value = 0
        self.imperial_value = 1

        self.units_lookup = {
            'metric': self.metric_value,
            'Metric': self.metric_value,
            'M': self.metric_value,
            'm': self.metric_value,
            'Imperial': self.imperial_value,
            'imperial': self.imperial_value,
            'I': self.imperial_value,
            'i': self.imperial_value
        }

        self.units_output = {
            self.metric_value: 'm',
            self.imperial_value: 'ft'
        }

        self.wind_units_output = {
            self.metric_value: 'm/s',
            self.imperial_value: 'ft/hr'
        }

        self.units_output_temp = {
            self.metric_value: degree_sign + 'C',
            self.imperial_value: degree_sign + 'C'
        }

        self.temp_lookup_value_C = 0
        self.temp_lookup_value_F = 1
        self.temp_lookup_value_K = 2
        self.temp_lookup_value_R = 3

        self.temperature_lookup = {
            'c': self.temp_lookup_value_C,
            'C': self.temp_lookup_value_C,
            degree_sign + 'c': self.temp_lookup_value_C,
            degree_sign + 'C': self.temp_lookup_value_C,
            'celsius': self.temp_lookup_value_C,
            'Celsius': self.temp_lookup_value_C,

            'f': self.temp_lookup_value_F,
            'F': self.temp_lookup_value_F,
            degree_sign + 'f': self.temp_lookup_value_F,
            degree_sign + 'F': self.temp_lookup_value_F,
            'fahrenheit': self.temp_lookup_value_F,
            'Fahrenheit': self.temp_lookup_value_F,

            'k': self.temp_lookup_value_K,
            'K': self.temp_lookup_value_K,
            degree_sign + 'k': self.temp_lookup_value_K,
            degree_sign + 'K': self.temp_lookup_value_K,
            'kelvin': self.temp_lookup_value_K,
            'Kelvin': self.temp_lookup_value_K,

            'r': self.temp_lookup_value_R,
            'R': self.temp_lookup_value_R,
            degree_sign + 'r': self.temp_lookup_value_R,
            degree_sign + 'R': self.temp_lookup_value_R,
            'rankine': self.temp_lookup_value_R,
            'Rankine': self.temp_lookup_value_R,
        }

        self.length_lookup_value_mm = 0
        self.length_lookup_value_cm = 1
        self.length_lookup_value_dm = 2
        self.length_lookup_value_m = 3
        self.length_lookup_value_mil = 4
        self.length_lookup_value_inch = 5
        self.length_lookup_value_foot = 6
        self.length_lookup_value_mile = 7

        self.length_lookup = {
            'mm': self.length_lookup_value_mm,
            'milli': self.length_lookup_value_mm,
            'millimeter': self.length_lookup_value_mm,
            'millimeters': self.length_lookup_value_mm,

            'cm': self.length_lookup_value_cm,
            'centi': self.length_lookup_value_cm,
            'centimeter': self.length_lookup_value_cm,
            'centimeters': self.length_lookup_value_cm,

            'dm': self.length_lookup_value_dm,
            'deci': self.length_lookup_value_dm,
            'decimeter': self.length_lookup_value_dm,
            'decimeters': self.length_lookup_value_dm,

            'm': self.length_lookup_value_m,
            'meter': self.length_lookup_value_m,
            'meters': self.length_lookup_value_m,

            'mil': self.length_lookup_value_mil,
            'mils': self.length_lookup_value_mil,
            'thou': self.length_lookup_value_mil,
            'thousand': self.length_lookup_value_mil,

            'in': self.length_lookup_value_inch,
            'inch': self.length_lookup_value_inch,
            'inches': self.length_lookup_value_inch,

            'ft': self.length_lookup_value_foot,
            'foot': self.length_lookup_value_foot,
            'feet': self.length_lookup_value_foot

        }

    def runTest(self):
        t0 = time.time()
        config_list = self.import_config(dir_config)
        conductor_list, conductor_spec = self.import_conductor(dir_conductor)
        df_config = self.select_config(config_list)
        df_conductor, df_spec = self.select_conductor(conductor_list, conductor_spec)
        df_adjusted = self.unit_conversion(df_conductor, df_spec, df_config)

        df = self.add_calc_columns(df_adjusted)
        n, e, load, df_out, dfn, dfe, dfl = self.c_reporting(df)
        print(f'Normal ratings {n}')
        print(f'Emergency ratings {e}')
        print(f'load dumping ratings {load}')
        print(dfn)
        print(dfe)
        # excel_sheets = ('Normal', 'Emergency', 'Load Dump', 'Results')
        # self.toExcel(n, e, load, df_out, excel_sheets)
        t1 = time.time()
        t = t1 - t0
        print(f'time to compute {t} seconds')
        return None

    @staticmethod
    def import_config(dir):
        # Imports list of configurations from database
        # Returns list of configurations
        df = pd.io.api.ExcelFile(dir, 'openpyxl')
        config_list = pd.read_excel(df, sheet_name='config')
        return config_list

    @staticmethod
    def import_conductor(dir):
        # Imports list of conductors from database
        # Returns list of conductors and list of conductor specs with Normal/Emergency temperature ratings
        df = pd.io.api.ExcelFile(dir, 'openpyxl')
        conductor_list = pd.read_excel(df, sheet_name='conductors')
        conductor_spec = pd.read_excel(df, sheet_name='conductor spec')
        conductor_list.sort_values('Metal OD', ascending=True, inplace=True)
        conductor_spec.sort_values('Conductor Spec', ascending=True, inplace=True)
        return conductor_list, conductor_spec

    def select_config(self, df_config):
        # Select configuration settings
        _response = None
        config = None
        _df = df_config['config name'].values
        _df_dict = _df.tolist()
        try:
            print('Select Configuration')
            for _pos, _text in enumerate(_df):
                print(f"{_pos + 1}: {_text}")
            _response = int(input("Selection? "))
            # _response = 1
            print(f'{_df_dict[_response - 1]}')
            config = df_config[df_config['config name'] == _df[_response - 1]].reset_index(drop=True)
            # config = df_config[df_config['config name'] == _df[_response - 1]]
        except KeyError:
            print(f'{_response} is not a valid selection')
            self.select_config(df_config)
        except IndexError:
            print('Please select a valid configuration')
            self.select_config(df_config)
        except TypeError:
            print('Please select a valid configuration')
            self.select_config(df_config)
        except ValueError:
            print('Please select a valid configuration')
            self.select_config(df_config)
        return config

    @staticmethod
    def select_conductor(df_conductor_list, df_spec):
        _config_name = None
        _conductor_spec = None
        _conductor_size = None
        _conductor_stranding = None
        _conductor_core_stranding = None
        conductor_data = None

        # TODO method works, but it is a little clumsy
        # Select conductor spec
        _df = df_conductor_list.drop_duplicates(['Conductor Spec'])
        _df = _df['Conductor Spec'].values
        for _pos, _text in enumerate(_df):
            print(f"{_pos + 1}: {_text}")
        _response = int(input("Selection?"))
        # _response = 1
        _conductor_spec = _df[_response - 1]
        print(_conductor_spec)

        # Select conductor size
        _df = df_conductor_list[df_conductor_list['Conductor Spec'] == _conductor_spec].drop_duplicates(['Size'])
        _df = _df['Size'].values
        for _pos, _text in enumerate(_df):
            print(f"{_pos + 1}: {_text}")
        _response = int(input("Selection?"))
        # _response = 1
        _conductor_size = _df[_response - 1]
        print(_conductor_size)

        # Depending on conductor spec, only sizing is required
        # Check to see if a single item exists
        _df = df_conductor_list[df_conductor_list['Conductor Spec'] == _conductor_spec]
        _df = _df[_df['Size'] == _conductor_size]

        if not _df.shape[0] == 1:
            # Select conductor stranding
            _df = _df.drop_duplicates(['Cond Strand'])
            _df.sort_values('Cond Strand', ascending=True, inplace=True)
            _df = _df['Cond Strand'].values
            for _pos, _text in enumerate(_df):
                print(f"{_pos + 1}: {_text}")
            _response = int(input("Selection?"))
            # _response = 4
            _conductor_stranding = _df[_response - 1]

            # Depending on conductor spec, only sizing is required
            # Check to see if a single item exists
            _df = df_conductor_list[df_conductor_list['Conductor Spec'] == _conductor_spec]
            _df = _df[_df['Size'] == _conductor_size]
            _df = _df[_df['Cond Strand'] == _conductor_stranding]
            if not _df.shape[0] == 1:
                _df = _df.drop_duplicates(['Core Strand'])
                _df.sort_values('Core Strand', ascending=True, inplace=True)
                _df = _df['Core Strand'].values
                for _pos, _text in enumerate(_df):
                    print(f"{_pos + 1}: {_text}")
                _response = int(input("Selection?"))
                print(_df[0])
                _conductor_core_stranding = _df[_response - 1]
                # _df = _df[_df['Core Strand'] == _conductor_core_stranding]

        _df = df_conductor_list  # Reset list to original excel import, could probably do this neater, but it works

        if _conductor_core_stranding is None:
            if _conductor_stranding is None:
                conductor_data = _df.loc[
                    (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size)]
            else:
                conductor_data = _df.loc[
                    (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size) & (
                            _df['Cond Strand'] == _conductor_stranding)]
        else:
            conductor_data = _df.loc[
                (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size) & (
                        _df['Cond Strand'] == _conductor_stranding) & (_df['Core Strand'] == _conductor_core_stranding)]

        _df_spec = pd.DataFrame()
        df_spec = df_spec.loc[df_spec['Conductor Spec'] == _conductor_spec].reset_index(drop=True)
        df_conductor = conductor_data.reset_index(drop=True)
        return df_conductor, df_spec

    # todo use names in found standard
    def add_calc_columns(self, df):
        data = ['qc heat loss', 'qc0', 'qc1', 'qc2', 'uf', 'kf', 'pf', 'Qse', 'theta', 'hc: solar altitude', 'delta',
                'omega', 'chi', 'qs heat gain', 'solar altitude correction factor', 'rating',
                'qr heat loss', 'day of year', 'k angle', 'solar azimuth constant', 'solar azimuth']
        df = pd.concat([df.reset_index(drop=True), pd.DataFrame(columns=data)], axis=1)
        return df

    def wind_units_adjustment(self, value, calculations_units, config_units):
        # Convert units provided in configuration file to calculation required values
        # Report units defines how the units are converted, i.e. metric (m/s) or imperial (ft/h)

        kph_2_mps = 1 / 3.6  # m/s
        fps_2_mps = 0.3048  # m/s
        fph_2_mps = 8.467e-5  # m/s
        mph_2_mps = 0.44704  # m/s
        kn_2_mps = 1 / 1.943844  # m/s

        mps_2_fph = 1 / 11811.024  # ft/hr
        kph_2_fph = 1 / 3280.84  # ft/hr
        fps_2_fph = 1 / 3600  # ft/hr
        mph_2_fph = 5280  # ft/hr
        kn_2_fph = 1 / 6076.12  # ft/hr

        _wind_speed_lookup_value_mps = 0  # meters per second
        _wind_speed_lookup_value_kmh = 1  # kilometers per hour
        _wind_speed_lookup_value_fps = 2  # feet per second
        _wind_speed_lookup_value_fph = 3  # feet per second
        _wind_speed_lookup_value_mph = 4  # miles per hour
        _wind_speed_lookup_value_knots = 5  # knots

        _wind_speed_lookup = {
            'mps': _wind_speed_lookup_value_mps,
            'meters/s': _wind_speed_lookup_value_mps,
            'meters/sec': _wind_speed_lookup_value_mps,
            'meters/second': _wind_speed_lookup_value_mps,
            'm/s': _wind_speed_lookup_value_mps,

            'kmh': _wind_speed_lookup_value_kmh,
            'kilometers/s': _wind_speed_lookup_value_kmh,
            'kilometers/sec': _wind_speed_lookup_value_mps,
            'kilometers/second': _wind_speed_lookup_value_mps,
            'km/s': _wind_speed_lookup_value_kmh,

            'fps': _wind_speed_lookup_value_fps,
            'feet/s': _wind_speed_lookup_value_fps,
            'feet/sec': _wind_speed_lookup_value_fps,
            'feet/second': _wind_speed_lookup_value_fps,
            'ft/s': _wind_speed_lookup_value_fps,

            'fph': _wind_speed_lookup_value_fph,
            'feet/h': _wind_speed_lookup_value_fph,
            'feet/hr': _wind_speed_lookup_value_fph,
            'feet/hour': _wind_speed_lookup_value_fph,
            'foot/h': _wind_speed_lookup_value_fph,
            'foot/hr': _wind_speed_lookup_value_fph,
            'foot/hour': _wind_speed_lookup_value_fph,
            'ft/hr': _wind_speed_lookup_value_fph,
            'ft/h': _wind_speed_lookup_value_fph,

            'mph': _wind_speed_lookup_value_mph,
            'm.p.h.': _wind_speed_lookup_value_mph,
            'MPH': _wind_speed_lookup_value_mph,
            'mi/hour': _wind_speed_lookup_value_mph,

            'kn': _wind_speed_lookup_value_knots,
            'kt': _wind_speed_lookup_value_knots,
            'knot': _wind_speed_lookup_value_knots,
            'knots': _wind_speed_lookup_value_knots,
        }

        try:
            if value <= 0:
                return [0, self.wind_units_output[self.units_lookup[calculations_units]]]
            else:
                if self.units_lookup[calculations_units] == self.metric_value:  # metric report
                    if _wind_speed_lookup[config_units] == _wind_speed_lookup_value_mps:
                        return [value, self.wind_units_output[self.units_lookup[calculations_units]]]  # meters per second
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_kmh:
                        return [value / kph_2_mps, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_fps:
                        return [value / fps_2_mps, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_fph:
                        return [value / fph_2_mps, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_mph:
                        return [value / mph_2_mps, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_knots:
                        return [value / kn_2_mps, self.wind_units_output[self.units_lookup[calculations_units]]]
                elif self.units_lookup[calculations_units] == self.imperial_value:  # imperial report
                    if _wind_speed_lookup[config_units] == _wind_speed_lookup_value_mps:
                        return [value / mps_2_fph, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_kmh:
                        return [value / kph_2_fph, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_fps:
                        return [value / fps_2_fph, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_fph:
                        return [value, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_mph:
                        return [value / mph_2_fph, self.wind_units_output[self.units_lookup[calculations_units]]]
                    elif _wind_speed_lookup[config_units] == _wind_speed_lookup_value_knots:
                        return [value / kn_2_fph, self.wind_units_output[self.units_lookup[calculations_units]]]
                else:
                    return [0, self.wind_units_output[self.units_lookup[calculations_units]]]
        except KeyError:
            return "Error", "Error"

    def wind_angle_units_adjustment(self, value, config_units):
        # Convert units provided in configuration file to calculation required values
        # Report units defines how the units are converted, i.e. metric (m/s) or imperial (ft/h)

        calculations_units = 'deg'

        wind_angle_lookup_value_deg = 0  # degrees
        wind_angle_lookup_value_rad = 1  # radians

        wind_angle_lookup = {
            'deg': wind_angle_lookup_value_deg,
            'degs': wind_angle_lookup_value_deg,
            'degree': wind_angle_lookup_value_deg,
            'degrees': wind_angle_lookup_value_deg,

            'rad': wind_angle_lookup_value_rad,
            'rads': wind_angle_lookup_value_rad,
            'radian': wind_angle_lookup_value_rad,
            'radians': wind_angle_lookup_value_rad,
        }
        try:
            if value <= 0:
                return [0, calculations_units]
            else:
                if wind_angle_lookup[calculations_units] == wind_angle_lookup_value_deg:  # angle in degrees
                    return value, calculations_units
                elif wind_angle_lookup[calculations_units] == wind_angle_lookup_value_rad:  # angle in radians
                    return np.degrees(value), calculations_units
                else:
                    return [0, calculations_units]
        except KeyError:
            return "Error", "Error"

    def temp_units_adjustment(self, value, config_units):
        try:
            if self.temperature_lookup[config_units] == self.temp_lookup_value_C:  # return Celsius
                return [value, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_F:  # Fahrenheit to Celsius
                return [(value - 32) * 5 / 9, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_K:  # Kelvin to Celsius
                return [value + 273.15, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_R:  # Rankine to Celsius
                return [(value - 491.67) * 5 / 9, self.units_output_temp[0]]
            else:
                return [0, self.units_output_temp[self.units_lookup[0]]]
        except KeyError:
            return "Error", "Error"

    def length_units_adjustment(self, length, input_units, unit_type):
        # Convert units provided in configuration file to calculation required values
        # Report units defines how the units are converted, i.e. metric (m/s) or imperial (ft/h)

        mm_2_m = 1 / 1000
        cm_2_m = 1 / 100
        dm_2_m = 1 / 10
        mil_2_m = 2.54E-5
        in_2_m = 0.0254
        foot_2_m = 0.3048
        mile_2_m = 1609.34

        mm_2_ft = 0.00328084
        cm_2_ft = 0.0328084
        dm_2_ft = 0.328084
        m_2_ft = 3.28084
        mil_2_ft = 1 / 12000
        in_2_ft = 1 / 12
        mile_2_ft = 5280

        try:
            if self.units_lookup[unit_type] == self.metric_value:  # metric report
                if self.length_lookup[input_units] == self.length_lookup_value_mm:
                    return [length * mm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_cm:
                    return [length * cm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_dm:
                    return [length * dm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_m:
                    return [length, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_mil:
                    return [length * mil_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_inch:
                    return [length * in_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_foot:
                    return [length * foot_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_mile:
                    return [length * mile_2_m, self.units_output[self.units_lookup[unit_type]]]
            elif self.units_lookup[unit_type] == self.imperial_value:  # imperial report
                if self.length_lookup[input_units] == self.length_lookup_value_mm:
                    return [length * mm_2_ft, self.units_output[self.units_lookup[unit_type]]]  # meters per second
                elif self.length_lookup[input_units] == self.length_lookup_value_cm:
                    return [length * cm_2_ft, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_dm:
                    return [length * dm_2_ft, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_m:
                    return [length * m_2_ft, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_mil:
                    return [length * mil_2_ft, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_inch:
                    return [length * in_2_ft, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_foot:
                    return [length, self.units_output[self.units_lookup[unit_type]]]
                elif self.length_lookup[input_units] == self.length_lookup_value_mile:
                    return [length * mile_2_ft, self.units_output[self.units_lookup[unit_type]]]
            else:
                return [0, self.units_output[self.units_lookup[unit_type]]]
        except KeyError:
            return "Error", "Error"

    def diameter_units_adjustment(self, length, input_units, output_units, unit_type):
        # todo fix this, doesn't work like it needs to
        # Convert units provided in configuration file to calculation required values
        # Report units defines how the units are converted, i.e. metric (m/s) or imperial (ft/h)

        mm_2_m = 1 / 1000
        cm_2_m = 1 / 100
        dm_2_m = 1 / 10
        mil_2_m = 2.54E-5
        in_2_m = 0.0254
        foot_2_m = 0.3048
        mile_2_m = 1609.34

        mm_2_in = 1 / 25.4
        cm_2_in = 1 / 2.54
        dm_2_in = 1 / .254
        m_2_in = 1 / 0.0254
        mil_2_in = 1 / 1000
        ft_2_in = 12
        mile_2_in = 63360

        length_lookup_value_mm = 0
        length_lookup_value_cm = 1
        length_lookup_value_dm = 2
        length_lookup_value_m = 3
        length_lookup_value_mil = 4
        length_lookup_value_inch = 5
        length_lookup_value_foot = 6
        length_lookup_value_mile = 7

        length_lookup = {
            'mm': length_lookup_value_mm,
            'milli': length_lookup_value_mm,
            'millimeter': length_lookup_value_mm,
            'millimeters': length_lookup_value_mm,

            'cm': length_lookup_value_cm,
            'centi': length_lookup_value_cm,
            'centimeter': length_lookup_value_cm,
            'centimeters': length_lookup_value_cm,

            'dm': length_lookup_value_dm,
            'deci': length_lookup_value_dm,
            'decimeter': length_lookup_value_dm,
            'decimeters': length_lookup_value_dm,

            'm': length_lookup_value_m,
            'meter': length_lookup_value_m,
            'meters': length_lookup_value_m,

            'mil': length_lookup_value_mil,
            'mils': length_lookup_value_mil,
            'thou': length_lookup_value_mil,
            'thousand': length_lookup_value_mil,

            'in': length_lookup_value_inch,
            'inch': length_lookup_value_inch,
            'inches': length_lookup_value_inch,

            'ft': length_lookup_value_foot,
            'foot': length_lookup_value_foot,
            'feet': length_lookup_value_foot

        }

        try:
            if self.units_lookup[unit_type] == self.metric_value:  # metric report
                if length_lookup[input_units] == length_lookup_value_mm:
                    return [length * mm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_cm:
                    return [length * cm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_dm:
                    return [length * dm_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_m:
                    return [length, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mil:
                    return [length * mil_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_inch:
                    return [length * in_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_foot:
                    return [length * foot_2_m, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mile:
                    return [length * mile_2_m, self.units_output[self.units_lookup[unit_type]]]
            elif self.units_lookup[unit_type] == self.imperial_value:  # imperial report
                if length_lookup[input_units] == length_lookup_value_mm:
                    return [length * mm_2_in, self.units_output[self.units_lookup[unit_type]]]  # meters per second
                elif length_lookup[input_units] == length_lookup_value_cm:
                    return [length * cm_2_in, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_dm:
                    return [length * dm_2_in, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_m:
                    return [length * m_2_in, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mil:
                    return [length * mil_2_in, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_inch:
                    return [length, self.units_output[self.units_lookup[unit_type]]]  # todo fix this
                elif length_lookup[input_units] == length_lookup_value_foot:
                    return [length, self.units_output[self.units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mile:
                    return [length * mile_2_in, self.units_output[self.units_lookup[unit_type]]]
            else:
                return [0, self.units_output[self.units_lookup[unit_type]]]
        except KeyError:
            return "Error", "Error"

    def unit_conversion(self, df_conductor, df_spec, df_config):
        # Todo remove other adjustment methods and move to UnitConversion.py
        unit_selection = None

        df_conductor_wind_list = pd.DataFrame()
        df_conductor_wind_angle_list = pd.DataFrame()
        df_conductor_length_list = pd.DataFrame()
        df_conductor_temp_list = pd.DataFrame()
        df_conductor_spec_temp_list = pd.DataFrame()
        df_config_temp_list = pd.DataFrame()
        df_config_length_list = pd.DataFrame()

        df_conductor_adjusted = pd.DataFrame(df_conductor)
        df_config_adjusted = pd.DataFrame(df_config)

        calculation_units = df_config.at[0, 'calculation units']

        if self.units_lookup[calculation_units] == self.metric_value:
            unit_selection = 2
        elif self.units_lookup[calculation_units] == self.imperial_value:
            unit_selection = 3

        # value, config unit, required unit for metric calculation, required unit for imperial calculation
        conductor_wind_list = (
            ('normal wind speed', 'normal wind speed units', 'm/s', 'ft/hr'),
            ('emergency wind speed', 'emergency wind speed units', 'm/s', 'ft/hr'),
        )
        conductor_wind_angle_list = (
            ('normal wind angle', 'normal wind angle units', 'deg', 'deg'),
            ('emergency wind angle', 'emergency wind angle units', 'deg', 'deg')
        )
        conductor_length_list = (
            ('Cond Wire Diameter', 'Cond Wire Diameter Units', 'mm', 'in'),
            ('Core Wire Diameter', 'Core Wire Diameter Units', 'mm', 'in'),
            ('Core OD', 'Core OD Units', 'mm', 'in'),
            ('Metal OD', 'Metal OD Units', 'mm', 'in'),
            ('resistance distance', 'resistance distance units', 'm', 'ft')
        )
        conductor_temp_list = (
            ('low resistance temperature', 'resistance temperature units', 'C', 'C'),
            ('high resistance temperature', 'resistance temperature units', 'Ohms', 'Ohms')
        )
        conductor_spec_temp_list = (
            ('normal temperature rating', 'normal temperature rating units', 'C', 'C'),
            ('emergency temperature rating', 'emergency temperature rating units', 'C', 'C')
        )
        config_temp_list = (
            ('ambient air temperature lower range', 'ambient air temperature units', 'C', 'C'),
            ('ambient air temperature upper range', 'ambient air temperature units', 'C', 'C'),
            ('temperature increment', 'ambient air temperature units', 'C', 'C'),
            ('ambient air temperature', 'ambient air temperature units', 'C', 'C')
        )
        config_length_list = (
            ('elevation', 'elevation units', 'm', 'ft'),
        )

        for x in conductor_wind_list:
            (value, units) = self.wind_units_adjustment(df_config.at[0, x[0]], calculation_units, df_config.at[0, x[1]])
            df_conductor_wind_list.at[0, x[0]] = value
            df_conductor_wind_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_wind_angle_list:
            (value, units) = self.wind_angle_units_adjustment(df_config.at[0, x[0]], df_config.at[0, x[1]])
            df_conductor_wind_angle_list.at[0, x[0]] = value
            df_conductor_wind_angle_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_length_list:
            value_new = uc.length_convert(df_conductor.at[0, x[0]], df_conductor.at[0, x[1]], x[unit_selection])
            df_conductor_length_list.at[0, x[0]] = value_new
            df_conductor_length_list.at[0, x[1]] = x[unit_selection]
            df_conductor_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_conductor_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_temp_list:
            (value, units) = self.temp_units_adjustment(df_conductor.at[0, x[0]], df_conductor.at[0, x[1]])
            df_conductor_temp_list.at[0, x[0]] = value
            df_conductor_temp_list.at[0, x[1]] = x[unit_selection]
            df_conductor_adjusted.drop(columns=x[0], axis=1, inplace=True)
            try:
                df_conductor_adjusted.drop(columns=x[1], axis=1, inplace=True)
            except:
                None

        for x in conductor_spec_temp_list:
            (value, units) = self.temp_units_adjustment(df_spec.at[0, x[0]], df_spec.at[0, x[1]])
            df_conductor_spec_temp_list.at[0, x[0]] = value
            df_conductor_spec_temp_list.at[0, x[1]] = x[unit_selection]

        for x in config_temp_list:
            (value, units) = self.temp_units_adjustment(df_config[x[0]].values[0], df_config.at[0, x[1]])
            df_config_temp_list.at[0, x[0]] = value
            df_config_temp_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            try:
                df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)
            except:
                None

        for x in config_length_list:
            value = uc.length_convert(df_config.at[0, x[0]], df_config.at[0, x[1]], x[unit_selection])
            df_config_length_list.at[0, x[0]] = value
            df_config_length_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)

        df_conductor_adjusted = pd.concat([df_conductor_adjusted, df_conductor_wind_list.reset_index(drop=True),
                                           df_conductor_wind_angle_list.reset_index(drop=True),
                                           df_conductor_length_list.reset_index(drop=True),
                                           df_conductor_temp_list.reset_index(drop=True),
                                           df_conductor_spec_temp_list.reset_index(drop=True),
                                           df_config_temp_list.reset_index(drop=True),
                                           df_config_length_list.reset_index(drop=True)], axis=1)

        df_adjusted = pd.concat(
            [df_conductor_adjusted.reset_index(drop=True), df_config_adjusted.reset_index(drop=True)], axis=1)

        df_adjusted = df_adjusted.fillna(0)

        return df_adjusted

    def c_steady_state(self, df, df_N=None, df_E=None, _idx=None):
        # Configuration setup
        conductor_projection = None
        calculation_units = df.at[_idx, 'calculation units']
        elevation = df.at[_idx, 'elevation']
        emissivity = df.at[_idx, 'emissivity']
        solar_absorptivity = df.at[_idx, 'solar absorptivity']
        atmosphere = df.at[_idx, 'atmosphere']
        latitude = df.at[_idx, 'latitude']
        day = df.at[_idx, 'day']
        month = df.at[_idx, 'month']
        year = df.at[_idx, 'year']
        hour = df.at[_idx, 'hour']
        conductor_direction = df.at[_idx, 'conductor direction']
        #
        ambient_air_temp = df.at[_idx, 'ambient air temperature']

        conductor_wind_normal = df.at[_idx, 'normal wind speed']
        conductor_wind_emergency = df.at[_idx, 'emergency wind speed']

        conductor_wind_angle_normal = df.at[_idx, 'normal wind angle']
        conductor_wind_angle_emergency = df.at[_idx, 'emergency wind angle']

        conductor_temperature = df.at[_idx, 'conductor temperature']

        # Conductor setup
        diameter = df.at[_idx, 'Metal OD']

        if self.units_lookup[calculation_units] == self.metric_value:
            conductor_projection = diameter / 1000
        elif self.units_lookup[calculation_units] == self.imperial_value:
            conductor_projection = diameter / 12

        d = {'high resistance Ω/unit': [df.at[0, 'high resistance Ω/unit']],  # high resistance
             'low resistance Ω/unit': [df['low resistance Ω/unit'].values[0]],  # low resistance
             'resistance temperature unit': [df['resistance temperature units'].values[0]],
             # resistance distance unit (C, F, K, R)
             'high resistance temperature': [df['high resistance temperature'].values[0]],
             # high resistance temp
             'low resistance temperature': [df['low resistance temperature'].values[0]],  # low resistance temp
             'resistance distance': [df['resistance distance'].values[0]],  # resistance distance
             'resistance distance unit': [df['resistance distance units'].values[0]]
             # resistance distance unit (mile/meter/etc...)
             }

        conductor_resistance = pd.DataFrame(d)
        # Normal Rating (wind speed and angle variable)

        normal_current_rating = self.c_SSRating(calculation_units, diameter, conductor_temperature,
                                         ambient_air_temp, elevation, conductor_wind_angle_normal,
                                         conductor_wind_normal, emissivity, solar_absorptivity, atmosphere,
                                         latitude, day, month,
                                         year, hour, conductor_direction, conductor_projection, conductor_resistance,
                                         df_N, _idx)

        # Emergency Rating (wind speed and angle variable)

        emergency_current_rating = self.c_SSRating(calculation_units, diameter, conductor_temperature,
                                            ambient_air_temp, elevation,
                                            conductor_wind_angle_emergency, conductor_wind_emergency, emissivity,
                                            solar_absorptivity,
                                            atmosphere, latitude,
                                            day, month, year, hour, conductor_direction, conductor_projection,
                                            conductor_resistance, df_E, _idx)

        return normal_current_rating, emergency_current_rating

    def c_load_dump(self, df, _idx):
        # Configuration setup
        conductor_projection = None
        _calculation_units = df.at[_idx, 'calculation units']
        _threshold = df.at[_idx, 'threshold']
        _elevation = df.at[_idx, 'elevation']
        _emissivity = df.at[_idx, 'emissivity']
        _solar_absorptivity = df.at[_idx, 'solar absorptivity']
        _atmosphere = df.at[_idx, 'atmosphere']
        _latitude = df.at[_idx, 'latitude']
        _day = df.at[_idx, 'day']
        _month = df.at[_idx, 'month']
        _year = df.at[_idx, 'year']
        _hour = df.at[_idx, 'hour']
        _conductor_direction = df.at[_idx, 'conductor direction']
        #
        _conductor_spec = df.at[_idx, 'Conductor Spec']
        _ambient_air_temp = df.at[_idx, 'ambient air temperature']

        _conductor_temp_normal = df.at[_idx, 'normal temperature rating']
        _conductor_temp_emergency = df.at[_idx, 'emergency temperature rating']
        _conductor_wind_emergency = df.at[_idx, 'emergency wind speed']
        conductor_temperature = df.at[_idx, 'conductor temperature']

        # Conductor setup
        _diameter = df.at[_idx, 'Metal OD']

        if self.units_lookup[_calculation_units] == self.metric_value:
            conductor_projection = _diameter / 1000
        elif self.units_lookup[_calculation_units] == self.imperial_value:
            conductor_projection = _diameter / 12

        d = {'high resistance Ω/unit': [df['high resistance Ω/unit'].values[0]],  # high resistance
             'low resistance Ω/unit': [df['low resistance Ω/unit'].values[0]],  # low resistance
             'resistance temperature unit': [df['resistance temperature units'].values[0]],
             # resistance distance unit (C, F, K, R)
             'high resistance temperature': [df['high resistance temperature'].values[0]],
             # high resistance temp
             'low resistance temperature': [df['low resistance temperature'].values[0]],  # low resistance temp
             'resistance distance': [df['resistance distance'].values[0]],  # resistance distance
             'resistance distance unit': [df['resistance distance units'].values[0]]
             # resistance distance unit (mile/meter/etc...)
             }

        _conductor_resistance = pd.DataFrame(d)

        _wind_angle = df.at[_idx, 'emergency wind angle']
        _Vw = df.at[_idx, 'emergency wind speed']

        # _al, _cu, _stl, _alw
        _mcp = self.c_mcp(df.at[_idx, 'Al Weight'] / 1000,
                          df.at[_idx, 'Cu Weight'] / 1000,
                          df.at[_idx, 'St Weight'] / 1000,
                          df.at[_idx, 'Alw Weight'] / 1000)

        _tau = 15
        _time = _tau * 60

        load_dump = self.load_dump(_threshold, _calculation_units, _diameter, _conductor_temp_normal,
                                   _conductor_temp_emergency, conductor_temperature,
                                   _ambient_air_temp,
                                   _elevation,
                                   _wind_angle, _conductor_wind_emergency, _emissivity, _solar_absorptivity,
                                   _atmosphere,
                                   _latitude,
                                   _day, _month, _year, _hour, _conductor_direction, conductor_projection,
                                   _conductor_resistance, _tau, _mcp)
        return load_dump
    @staticmethod
    def c_temperature_range(df_adjusted):
        ambient_lower_range = df_adjusted.at[0, 'ambient air temperature lower range']
        ambient_upper_range = df_adjusted.at[0, 'ambient air temperature upper range']
        ambient_temp_inc = df_adjusted.at[0, 'temperature increment']
        conductor_normal_rating = df_adjusted.at[0, 'normal temperature rating']
        conductor_emergency_rating = df_adjusted.at[0, 'emergency temperature rating']
        conductor_temp_steps = 6

        if conductor_normal_rating == conductor_emergency_rating:
            temp_range_conductor = np.arange(conductor_normal_rating - conductor_temp_steps * ambient_temp_inc,
                                             conductor_normal_rating + ambient_temp_inc, ambient_temp_inc)
        else:
            temp_range_conductor = np.arange(conductor_normal_rating - conductor_temp_steps * ambient_temp_inc,
                                             conductor_emergency_rating + ambient_temp_inc, ambient_temp_inc)

        temp_range_ambient = np.arange(ambient_lower_range, ambient_upper_range + ambient_temp_inc, ambient_temp_inc)

        return temp_range_ambient, temp_range_conductor


    def c_reporting(self, df_adjusted):
        _idx = 0
        temp_range_ambient, temp_range_conductor = self.c_temperature_range(df_adjusted)
        _temp_list = (
            ('ambient air temperature', 'ambient air temperature units'),
            ('normal temperature rating', 'normal temperature rating units'),
            ('emergency temperature rating', 'emergency temperature rating units')
        )

        normal_report = np.empty((temp_range_conductor.shape[0], 0))
        emergency_report = np.empty((temp_range_conductor.shape[0], 0))
        load_report = np.empty((temp_range_conductor.shape[0], 0))

        df = pd.DataFrame(df_adjusted)
        df_N = pd.DataFrame(df)
        df_E = pd.DataFrame(df)
        df_L = pd.DataFrame(df)

        for i, element_i in enumerate(temp_range_ambient):
            normal_holder = np.empty((temp_range_conductor.shape[0], 1))
            normal_holder[:] = np.nan
            emergency_holder = np.empty((temp_range_conductor.shape[0], 1))
            emergency_holder[:] = np.nan
            load_holder = np.empty((temp_range_conductor.shape[0], 1))
            load_holder[:] = np.nan
            for j, element_j in enumerate(temp_range_conductor):
                df = pd.concat([df, df.iloc[[0]]], axis=0, ignore_index=True)
                df.at[_idx, 'ambient air temperature'] = element_i
                df.at[_idx, 'conductor temperature'] = element_j
                normal_rating, emergency_rating = self.c_steady_state(df, df_N, df_E, _idx=_idx)
                if j == 0:  # todo make this better
                    ppp = 1
                    # load_rating = self.c_load_dump(df, _idx)
                    load_rating = 0
                else:
                    load_rating = load_holder[0]
                np.put(normal_holder, j, normal_rating)
                np.put(emergency_holder, j, emergency_rating)
                np.put(load_holder, j, load_rating)
                _idx = _idx + 1
            normal_report = np.append(normal_report, normal_holder, axis=1)
            emergency_report = np.append(arr=emergency_report, values=emergency_holder, axis=1)
            load_report = np.append(arr=load_report, values=load_holder, axis=1)

        df_normal = pd.DataFrame(data=normal_report, columns=temp_range_ambient, index=temp_range_conductor)
        df_emergency = pd.DataFrame(data=emergency_report, columns=temp_range_ambient,
                                    index=temp_range_conductor)
        df_load = pd.DataFrame(data=load_report, columns=temp_range_ambient, index=temp_range_conductor)

        df_normal.index.name = 'Conductor Temperature'
        df_emergency.index.name = 'Conductor Temperature'
        df_load.index.name = 'Conductor Temperature'

        return df_normal, df_emergency, df_load, df, df_N, df_E, df_L

    def toExcel(self, df_n, df_e, df_l, df_output, sheetnames):
        # todo add some note to the DF that shows which calculation it is day/night normal/emergency/load dump
        with pd.ExcelWriter("test.xlsx") as writer:
            df_n.to_excel(writer, sheet_name=sheetnames[0])
            df_e.to_excel(writer, sheet_name=sheetnames[1])
            df_l.to_excel(writer, sheet_name=sheetnames[2])
            df_output.to_excel(writer, sheet_name=sheetnames[3])
        return None

    # Functions
    @staticmethod
    def current_steady_state(qr, qs, qc, r):
        """
        Calculates steady state current rating and returns results (Amps)
        :param qr: Radiated heat loss (W/m or W/ft)
        :param qs: Heat gain from The Sun (W/m or W/ft)
        :param qc: Convected heat loss (W/m or W/ft)
        :param r: Conductor resistance (ohms)
        :return: Steady state current rating (Amps)
        """
        rating = np.sqrt((qr - qs + qc) / r)
        return rating

    @staticmethod
    def c_uf(units, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates dynamic viscosity of air and returns results (Pa-s or lb/ft-hr)
        :param units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Dynamic viscosity of air (Pa-s or lb/ft-hr)
        """
        t_film = (conductor_temp + ambient_air_temp) / 2
        # Todo update unit detection method
        if units == 'metric':
            uf = (1.458 * 10 ** -6 * (t_film + 273.15) ** 1.5) / (t_film + 383.4)
        else:
            uf = (0.00353 * (t_film + 273.15) ** 1.5) / (t_film + 383.4)
        if df is not None:
            df.at[_idx, 'uf'] = uf
        return uf

    @staticmethod
    def c_kf(units, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates thermal conductivity of air and returns results (W/m*C or W/ft*C)
        :param units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Thermal conductivity of air (W/m*C or W/ft*C)
        """
        t_film = (conductor_temp + ambient_air_temp) / 2
        # Todo update unit detection method
        if units == 'metric':
            kf = 2.424 * 10 ** -2 + 7.477 * 10 ** -5 * t_film - 4.407 * 10 ** -9 * t_film ** 2
        else:
            kf = 0.007388 + 2.279 * 10 ** -5 * t_film - 1.343 * 10 ** -9 * t_film ** 2
        if df is not None:
            df.at[_idx, 'kf'] = kf
        return kf

    @staticmethod
    def c_pf( units, conductor_temp, ambient_air_temp, elevation, df=None, _idx=0):
        """
        Calculates air density and returns results (kg/m^3 or lb/ft^3)
        :param units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param elevation: elevation of conductors
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Air density (kg/m^3 or lb/ft^3
        """
        t_film = (conductor_temp + ambient_air_temp) / 2
        if units == 'metric':
            pf = (1.293 - 1.525 * 10 ** -4 * elevation + 6.379 * 10 ** -9 * elevation ** 2) / (
                    1 + 0.00367 * t_film)
        else:
            pf = (0.080695 - (2.901 * 10 ** -6) * elevation + (3.7 * 10 ** -11) * (elevation ** 2)) / (
                    1 + 0.00367 * t_film)
        if df is not None:
            df.at[_idx, 'pf'] = pf
        return pf

    @staticmethod
    def c_cond_resistance(conductor_temp, conductor_resistance, df=None, _idx=0):
        """
        Calculates resistance of conductor and returns results (ohm)
        :param conductor_temp: Conductor temperature (C)
        :param conductor_resistance: Dataframe containing resistance values/temperatures/distance
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Resistance (ohm)
        """
        high_resistance_ohm_per_unit_distance = conductor_resistance.at[0, 'high resistance Ω/unit']
        low_resistance_ohm_per_unit_distance = conductor_resistance.at[0, 'low resistance Ω/unit']
        high_resistance_temperature = conductor_resistance.at[0, 'high resistance temperature']
        low_resistance_temperature = conductor_resistance.at[0, 'low resistance temperature']
        resistance_distance = conductor_resistance.at[0, 'resistance distance']
        high_resistance = high_resistance_ohm_per_unit_distance / resistance_distance
        low_resistance = low_resistance_ohm_per_unit_distance / resistance_distance

        # todo add mention in report that conductor temperature is higher than resistance temperature and the results
        #  are less conservative
        #  overall resistance at temperature might be lower than what actually occurs physically
        #  reference page 10 738-2006
        resistance = (
                ((high_resistance - low_resistance) /
                 (high_resistance_temperature - low_resistance_temperature)) *
                (conductor_temp - low_resistance_temperature) + low_resistance
        )
        if df is not None:
            df.at[_idx, 'conductor resistance'] = resistance

        return resistance

    def c_Qs(self, units,  atmosphere, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates total solar and sky radiated heat flux rate (W/m^2) and returns results
        :param units: Units: 'Metric' or 'Imperial'
        :param atmosphere: Atmospheric conditions 'Industrial' or 'Clear'
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Total solar and sky radiated heat flux rate (W/m^2)
        """

        #Todo match unit detection method
        if units == 'metric':
            if atmosphere == 'clear':
                # Clear atmosphere
                aa = -42.2391
                bb = 63.8044
                cc = -1.9220
                dd = 3.46921E-2
                ee = -3.61118E-4
                ff = 1.94318E-6
                gg = -4.07608E-9
            else:
                # Industrial atmosphere
                aa = 53.1821
                bb = 14.2110
                cc = 6.6138E-1
                dd = -3.1658E-2
                ee = 5.4654E-4
                ff = -4.3446E-6
                gg = 1.3236E-8
        else:
            if atmosphere == 'clear':
                # Clear atmosphere
                aa = -3.9241
                bb = 5.9276
                cc = -1.7856E-1
                dd = 3.223E-3
                ee = -3.3549E-5
                ff = 1.8053E-7
                gg = -3.7868E-10
            else:
                # Industrial atmosphere
                aa = 4.9408
                bb = 1.3202
                cc = 6.1444E-2
                dd = -2.9411E-3
                ee = 5.07752E-5
                ff = -4.03627E-7
                gg = 1.22967E-9

        solar_altitude = self.c_solar_altitude(latitude, day, month, year, hour, df, _idx)
        radiated_heat_flux_rate = aa + bb * solar_altitude + cc * solar_altitude ** 2 + dd * solar_altitude ** 3 + ee * solar_altitude ** 4 + ff * solar_altitude ** 5 + gg * solar_altitude ** 6

        if df is not None:
            df.at[_idx, 'Qs'] = radiated_heat_flux_rate

        return radiated_heat_flux_rate

    def c_Qse(self, units, elevation, atmosphere, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates elevation corrected total solar and sky radiated heat flux rate (W/m^2) and returns results
        :param units: Units: 'Metric' or 'Imperial'
        :param elevation: elevation of conductors
        :param atmosphere: Atmospheric conditions 'Industrial' or 'Clear'
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: elevation corrected total solar and sky radiated heat flux rate (W/m^2)
        """
        ksolar = self.c_ksolar(units, elevation, df, _idx)
        Qs = self.c_Qs(units, atmosphere, latitude, day, month, year, hour, df, _idx)
        Qse = ksolar * Qs

        if df is not None:
            df.at[_idx, 'Qse'] = Qse

        return Qse

    @staticmethod
    def c_day_of_year(day, month, year, df, _idx):
        """
        Calculates day of year and returns results (int), ex: Jan 21 day_of_year = 21, Feb 12 day_of_year = 43
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Day of year (int)
        """
        error = False
        day_of_year = 0
        try:
            date = datetime.datetime(int(year), int(month), int(day))
            day_of_year = int(date.strftime("%j"))  # Get the day of the year
        except ValueError:  # if date is not valid
            df.at[_idx, 'Error'] = df.at[_idx, 'Error'].astype(
                str) + ' c_day_of_year: Check day/month/year, using 06/10/2009 '
            date = datetime.datetime(2009, 6, 10)
            day_of_year = int(date.strftime("%j"))  # Get the day of the year
            df.at[_idx, 'day of year'] = day_of_year
            error = True

        if df is not None and not error:
            df.at[_idx, 'day of year'] = day_of_year
        return day_of_year

    @staticmethod
    def c_ksolar(units, elevation, df=None, _idx=0):
        """
        Calculates solar heat multiplying factor (kSolar) and returns results
        :param units: Units: 'Metric' or 'Imperial'
        :param elevation: elevation of conductors
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar Heat multiplying factor
        """
        aks = 1.0
        bks = 3.500E-5
        cks = -1.000E-9

        # Todo update unit detection to match other methods used
        if units == 'metric':
            # meters
            if elevation < 1000:
                solar_heat_factor = 1.0
            elif 1000 <= elevation < 2000:
                solar_heat_factor = 1.10
            elif 2000 <= elevation < 4000:
                solar_heat_factor = 1.19
            elif 4000 <= elevation:
                solar_heat_factor = 1.28
        else:
            # feet
            if elevation < 5000:
                solar_heat_factor = 1.0
            elif 5000 <= elevation < 10000:
                solar_heat_factor = 1.15
            elif 10000 <= elevation < 15000:
                solar_heat_factor = 1.25
            elif 15000 <= elevation:
                solar_heat_factor = 1.30

        solar_altitude_correction_factor = aks + bks * elevation + cks * elevation ** 2

        if df is not None:
            df.at[_idx, 'solar altitude correction factor'] = solar_heat_factor

        return solar_altitude_correction_factor

    @staticmethod
    def c_chi(omega, latitude, delta, df=None, _idx=0):
        """
        Calculates solar azimuth variable returns value (no units), used to calculate solar azimuth
        :param omega: hour angle (radians)
        :param latitude: latitude
        :param delta: solar declination (radians)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar Azimuth variable (no units)
        """
        chi = np.sin(omega) / (np.sin(latitude) * np.cos(omega) - np.cos(latitude) * np.tan(delta))
        if df is not None:
            df.at[_idx, 'chi'] = chi
        return chi

    @staticmethod
    def c_delta(day_of_year, df=None, _idx=0):
        """
        Calculates solar declination and returns results (radians)
        :param day_of_year: Day of Year, ex: Jan 21 day_of_year = 21, Feb 12 day_of_year = 43
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar declination angle (radians)
        """
        delta = np.radians(23.4583 * np.sin(np.radians((284 + day_of_year) / 365 * 360)))
        if df is not None:
            df.at[_idx, 'delta'] = delta
        return delta

    @staticmethod
    def c_omega(hour, df=None, _idx=0):
        """
        Calculates hour angle and returns results (radians)
        :param hour: Hour of day (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: hour angle (radians)
        """
        omega = np.radians((hour / 100 - 12) * 15)
        if df is not None:
            df.at[_idx, 'omega'] = omega
        return omega

    @staticmethod
    def c_solar_constant(omega, chi, df=None, _idx=0):
        """
        Calculates solar azimuth constant (C) and returns results (degrees)
        :param omega: hour angle (radians)
        :param chi: Solar Azimuth variable (no units)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar azimuth constant (degrees)
        """
        omega = np.degrees(omega)
        if -180 <= omega < 0:
            if chi >= 0:
                solar_azimuth_constant = 0
            else:
                solar_azimuth_constant = 180
        else:
            if chi >= 0:
                solar_azimuth_constant = 180
            else:
                solar_azimuth_constant = 360
        if df is not None:
            df.at[_idx, 'solar azimuth constant'] = solar_azimuth_constant
        return solar_azimuth_constant

    def c_solar_azimuth(self, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates solar azimuth and returns results (degrees)
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar Azimuth (degrees)
        """
        latitude = np.radians(latitude)
        day_of_year = self.c_day_of_year(day, month, year, df, _idx)
        delta = self.c_delta(day_of_year, df, _idx)
        omega = self.c_omega(hour, df, _idx)
        chi = self.c_chi(omega, latitude, delta, df, _idx)
        solar_constant = self.c_solar_constant(omega, chi, df, _idx)
        solar_azimuth = solar_constant + np.degrees(np.arctan(chi))
        if df is not None:
            df.at[_idx, 'solar azimuth'] = solar_azimuth
        return solar_azimuth

    def c_solar_altitude(self, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates solar altitude (Hc) and returns results (degrees)
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar altitude (degrees)
        """
        latitude = np.radians(latitude)
        day_of_year = self.c_day_of_year(day, month, year, df, _idx)
        delta = self.c_delta(day_of_year, df, _idx)
        omega = self.c_omega(hour, df, _idx)
        solar_altitude = np.degrees(np.arcsin(np.cos(latitude) * np.cos(delta) * np.cos(omega) +
                                              np.sin(latitude) * np.sin(delta)))
        if df is not None:
            df.at[_idx, 'hc: solar altitude'] = solar_altitude
        return solar_altitude

    def c_Theta(self, latitude, day, month, year, hour, conductor_direction, df=None, _idx=0):
        """
        Calculates Effective angles of incidence of the Sun's rays (θ) and returns results(degrees)
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param conductor_direction: Direction conductors run 'North/South' or 'East/West' (string)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Effective angles of incidence of the Sun's rays (degrees)
        """
        direction_lookup_value_ns = 1
        direction_lookup_value_ew = 2

        direction_lookup = {
            'n/s': direction_lookup_value_ns,
            'n-s': direction_lookup_value_ns,
            'north/south': direction_lookup_value_ns,
            'north-south': direction_lookup_value_ns,
            's/n': direction_lookup_value_ns,
            's-n': direction_lookup_value_ns,
            'south/north': direction_lookup_value_ns,
            'north-south': direction_lookup_value_ns,

            'e/w': direction_lookup_value_ew,
            'e-w': direction_lookup_value_ew,
            'east/west': direction_lookup_value_ew,
            'east-west': direction_lookup_value_ew,
            'w/e': direction_lookup_value_ew,
            'w-e': direction_lookup_value_ew,
            'west/east': direction_lookup_value_ew,
            'west-east': direction_lookup_value_ew,
        }

        if direction_lookup[conductor_direction.lower()] == direction_lookup_value_ns:
            _Z1 = 90
        else:
            _Z1 = 0

        solar_altitude = self.c_solar_altitude(latitude, day, month, year, hour, df, _idx)
        solar_azimuth = self.c_solar_azimuth(latitude, day, month, year, hour, df, _idx)

        theta = np.degrees(np.arccos(np.cos(np.radians(solar_altitude)) * np.cos(np.radians(solar_azimuth - _Z1))))

        if df is not None:
            df.at[_idx, 'theta'] = theta

        return theta

    @staticmethod
    def c_k_angle(wind_angle, df=None, _idx=0):
        """
        Calculates wind direction factor and returns results (no units)
        :param wind_angle: Angle between conductor and applied wind (degrees)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Wind direction factor (no units)
        """
        k_angle = 1.194 - np.cos(np.radians(wind_angle)) + 0.194 * np.cos(np.radians(2 * wind_angle)) + 0.368 * \
                  np.sin(np.radians(2 * wind_angle))
        if df is not None:
            df.at[_idx, 'k angle'] = k_angle
        return k_angle

    def c_qsHeatGain(self, units, solar_absorptivity, elevation, atmosphere, latitude, day, month,
                     year, hour, conductor_direction, conductor_projection, df=None, _idx=0):
        """
        Calculates heat gain rate from the sun and returns results (W/m or W/ft)
        :param units: Units: 'Metric' or 'Imperial'
        :param solar_absorptivity: Solar absorptivity
        :param elevation: elevation of conductors
        :param atmosphere: Atmospheric conditions 'Industrial' or 'Clear'
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param conductor_direction: Direction conductors run 'North/South' or 'East/West' (string)
        :param conductor_projection: Projected area of the conductor per unit length (diameter / 1000 or diameter / 12)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return:
        """
        Qse = self.c_Qse(units, elevation, atmosphere, latitude, day, month, year, hour, df, _idx)
        theta = self.c_Theta(latitude, day, month, year, hour, conductor_direction, df, _idx)
        qs_heat_gain = solar_absorptivity * Qse * np.sin(np.radians(theta)) * conductor_projection

        if df is not None:
            df.at[_idx, 'qs heat gain'] = qs_heat_gain
        return qs_heat_gain

    @staticmethod
    def c_qrHeatLoss(units, diameter, emissivity, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates radiated heat loss rate per unit length and returns results (W/m or W/ft)
        :param units: Units: 'Metric' or 'Imperial'
        :param diameter: Conductor diameter (mm or in)
        :param emissivity: Emissivity
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Radiated heat loss rate per unit length (W/m or W/ft)
        """
        # Todo update unit detection method to match other areas
        if units == 'metric':
            # W/meters
            qr = 0.0178 * diameter * emissivity * (
                    ((conductor_temp + 273.15) / 100) ** 4 - ((ambient_air_temp + 273.15) / 100) ** 4)
        else:
            # W/feet
            qr = 0.138 * diameter * emissivity * (
                    ((conductor_temp + 273.15) / 100) ** 4 - ((ambient_air_temp + 273.15) / 100) ** 4)

        if df is not None:
            df.at[_idx, 'qr heat loss'] = qr
        return qr

    def c_qcHeatLoss(self, units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                     wind_speed, df=None, _idx=0):
        """
        Calculates convected heat loss rate per unit length and returns results (W/m or W/ft)
        :param units: Units: 'Metric' or 'Imperial'
        :param diameter: Conductor diameter (mm or in)
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param elevation: elevation of conductors
        :param wind_angle: Angle between conductor and applied wind (degrees)
        :param wind_speed: Wind speed (m/s or ft/hr)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Convected heat loss (W/m or W/ft)
        """
        k_angle = self.c_k_angle(wind_angle, df, _idx)
        uf = self.c_uf(units, conductor_temp, ambient_air_temp, df, _idx)
        kf = self.c_kf(units, conductor_temp, ambient_air_temp, df, _idx)
        pf = self.c_pf(units, conductor_temp, ambient_air_temp, elevation, df, _idx)

        # Todo update unit detection method to match elsewhere
        if units == 'metric':
            # W/meter
            # natural convection
            qc0 = 0.0205 * pf ** 0.5 * diameter ** 0.75 * (conductor_temp - ambient_air_temp) ** 1.25

            # low wind speeds
            qc1 = (1.01 + 0.0372 * ((diameter * pf * wind_speed) / uf) ** 0.52) * kf * k_angle * (
                    conductor_temp - ambient_air_temp)

            # high wind speeds
            qc2 = (0.0119 * ((diameter * pf * wind_speed) / uf) ** 0.6) * kf * k_angle * (
                    conductor_temp - ambient_air_temp)
            qc_heat_loss = np.amax((qc0, qc1, qc2))
        else:
            # W/feet

            # natural convection
            qc0 = 0.283 * pf ** 0.5 * diameter ** 0.75 * (conductor_temp - ambient_air_temp) ** 1.25

            # low wind speeds
            qc1 = (1.01 + 0.371 * ((diameter * pf * wind_speed) / uf) ** 0.52) * kf * k_angle * (
                    conductor_temp - ambient_air_temp)

            # high wind speeds
            qc2 = (0.1695 * ((diameter * pf * wind_speed) / uf) ** 0.6) * kf * k_angle * (
                    conductor_temp - ambient_air_temp)
            qc_heat_loss = np.amax((qc0, qc1, qc2))

        if df is not None:
            df.at[_idx, 'qc0'] = qc0
            df.at[_idx, 'qc1'] = qc1
            df.at[_idx, 'qc2'] = qc2
            df.at[_idx, 'qc heat loss'] = qc_heat_loss
        return qc_heat_loss

    def c_SSRating(self, units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                   wind_speed, emissivity, solar_absorptivity, atmosphere, latitude, day, month, year, hour,
                   conductor_direction, conductor_projection, conductor_resistance, df=None, _idx=0):
        """
        Calculates steady state current and returns results (Amps)
        :param units: Units: 'Metric' or 'Imperial'
        :param diameter: Conductor diameter (mm or in)
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param elevation: elevation of conductors
        :param wind_angle: Angle between conductor and applied wind (degrees)
        :param wind_speed: Wind speed (m/s or ft/hr)
        :param emissivity: Emissivity
        :param solar_absorptivity: Solar absorptivity
        :param atmosphere: Atmospheric conditions 'Industrial' or 'Clear'
        :param latitude: latitude
        :param day: Day of month (int)
        :param month: Month (int)
        :param year: Year (int)
        :param hour: Hour of day (int)
        :param conductor_direction: Direction conductors run 'North/South' or 'East/West' (string)
        :param conductor_projection: Projected area of the conductor per unit length (diameter / 1000 or diameter / 12)
        :param conductor_resistance:
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Steady state current (Amps)
        """
        qc = self.c_qcHeatLoss(units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                               wind_speed, df, _idx)
        qr = self.c_qrHeatLoss(units, diameter, emissivity, conductor_temp, ambient_air_temp, df, _idx)
        qs = self.c_qsHeatGain(units, solar_absorptivity, elevation, atmosphere, latitude, day, month, year,
                               hour, conductor_direction, conductor_projection, df, _idx)
        r_cond = self.c_cond_resistance(conductor_temp, conductor_resistance, df, _idx)
        rating = self.current_steady_state(qr, qs, qc, r_cond)

        if df is not None:
            df.at[_idx, 'rating'] = rating
        return rating

    @staticmethod
    def c_mcp(_al, _cu, _stl, _alw):
        _results = 433 * _al + 192 * _cu + 216 * _stl + 242 * _alw
        return _results

    def iTemp(self, _threshold, _units, _diameter, _conductor_temp_normal, _conductor_temp_emergency, _ambient_air_temp,
              _elevation, _wind_angle, _Vw,
              _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
              _conductor_direction, conductor_projection, _conductor_resistance):
        _max_iterations = 50
        _max = False
        _int = 0
        _ii = self.c_SSRating(_units, _diameter, _conductor_temp_normal, _ambient_air_temp, _elevation, _wind_angle, 0,
                              _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                              _conductor_direction, conductor_projection,
                              _conductor_resistance)

        lower_t = _ambient_air_temp  # conductor cannot be lower than ambient unless actively cooled
        upper_t = _conductor_temp_emergency * 3  #
        solve_t = (lower_t + upper_t) / 2
        threshold = _ii - self.c_SSRating(_units, _diameter, solve_t, _ambient_air_temp, _elevation, _wind_angle, _Vw,
                                          _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month,
                                          _year, _hour, _conductor_direction, conductor_projection,
                                          _conductor_resistance)
        if debug:
            print(f'Threshold: {threshold}')

        while np.abs(threshold) >= np.abs(_threshold) and not _max:
            if debug:
                print(f'range is: {lower_t}  ----  {solve_t}   ----   {upper_t}')
            if threshold < 0:
                upper_t = solve_t
                solve_t = (lower_t + upper_t) / 2
                if debug:
                    print(f'< {solve_t}')
            elif threshold > 0:
                lower_t = solve_t
                solve_t = (lower_t + upper_t) / 2
                if debug:
                    print(f'> {solve_t}')
            holder = self.c_SSRating(_units, _diameter, solve_t, _ambient_air_temp, _elevation, _wind_angle, _Vw,
                                     _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year,
                                     _hour, _conductor_direction, conductor_projection, _conductor_resistance)
            threshold = _ii - holder
            _int += 1
            if _int >= _max_iterations:
                _max = True
            if debug:
                print(f'Holder: {holder} Threshold: {threshold}')
                input(f'Iteration #: {_int} Press Enter to continue...')

        if _max:
            print('Unable to converge')
            print(f'Final result: Threshold: {threshold}....Solved input: {solve_t}....Number of Iterations {_int}')
            _results = solve_t
        else:
            _results = solve_t
        if debug:
            print(f'Max iterations: {_int}, Final result: Threshold: {threshold}....Solved input: {solve_t}')
            print("Initial Conductor Temp", _results)

        return _results

    def final_temp(self, _initial_temperature, _mcp, _conductor_resistance, _final_current, _initial_current, _tau):
        _r = self.c_cond_resistance(_initial_temperature, _conductor_resistance)
        results = (60 * _tau * _r * (_final_current ** 2 - _initial_current ** 2)) / _mcp + _initial_temperature
        return results

    def temp_conductor(self, _initial_temperature, _final_temperature, _time, _calc_tau):
        results = _initial_temperature + (_final_temperature - _initial_temperature) * (1 - np.exp(-_time / _calc_tau))
        return results

    def load_dump(self, _threshold, _calculation_units, _diameter, _conductor_temp_normal, _conductor_temp_emergency,
                  _conductor_temp, _ambient_air_temp, _elevation,
                  _wind_angle, _conductor_wind_emergency_adjusted,
                  _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                  _conductor_direction, conductor_projection, _conductor_resistance, _tau, _mcp):

        # Initial current, no wind
        # initial temperature, --> initial current with emergency wind applied
        # verify wind angles and naming conventions throughout mix of Vw and normal/emergency

        _max_iterations = 50
        _max = False
        _int = 0
        _initial_temperature = self.iTemp(_threshold, _calculation_units, _diameter, _conductor_temp_normal,
                                          _conductor_temp_emergency,
                                          _ambient_air_temp,
                                          _elevation,
                                          _wind_angle, _conductor_wind_emergency_adjusted, _emissivity,
                                          _solar_absorptivity, _atmosphere,
                                          _latitude,
                                          _day, _month, _year, _hour, _conductor_direction, conductor_projection,
                                          _conductor_resistance)

        _final_temperature = _conductor_temp_emergency / 0.632

        # calculate initial current

        _initial_current = self.c_SSRating(_calculation_units, _diameter, _conductor_temp_normal, _ambient_air_temp,
                                           _elevation,
                                           _wind_angle, 0,
                                           _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month,
                                           _year, _hour,
                                           _conductor_direction, conductor_projection, _conductor_resistance)

        _final_current = self.c_SSRating(_calculation_units, _diameter, _final_temperature, _ambient_air_temp,
                                         _elevation,
                                         _wind_angle, _conductor_wind_emergency_adjusted, _emissivity,
                                         _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                                         _conductor_direction, conductor_projection, _conductor_resistance)

        # _r = self.c_cond_resistance((_initial_temperature+_final_temperature)/2, _conductor_resistance)
        # todo pick a method, average/low/Ti
        _r = self.c_cond_resistance(_initial_temperature, _conductor_resistance)
        _calc_tau = (_mcp * (_final_temperature - _initial_temperature)) / (
                _r * (_final_current ** 2 - _initial_current ** 2)) / 60

        _tc = self.temp_conductor(_initial_temperature, _final_temperature, 15, _calc_tau)

        lower_t = _initial_temperature
        upper_t = _conductor_temp_emergency * 3
        solve_t = (lower_t + upper_t) / 2

        threshold = _conductor_temp_emergency - _tc
        if debug:
            print(f'Threshold: {threshold}')

        while np.abs(threshold) >= np.abs(_threshold) and not _max:
            if debug:
                print(f'range is: {lower_t}  ----  {solve_t}   ----   {upper_t}')
            if threshold < 0:
                upper_t = solve_t
                solve_t = (lower_t + upper_t) / 2
                if debug:
                    print(f'< {solve_t}')
            elif threshold > 0:
                lower_t = solve_t
                solve_t = (lower_t + upper_t) / 2
                if debug:
                    print(f'> {solve_t}')
                if not self.true_to_standard:
                    _r = self.c_cond_resistance(solve_t, _conductor_resistance)
            _final_current = self.c_SSRating(_calculation_units, _diameter, solve_t, _ambient_air_temp, _elevation,
                                             _wind_angle,
                                             _conductor_wind_emergency_adjusted,
                                             _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month,
                                             _year,
                                             _hour, _conductor_direction, conductor_projection, _conductor_resistance)
            _calc_tau = (_mcp * (solve_t - _initial_temperature)) / (
                    _r * (_final_current ** 2 - _initial_current ** 2)) / 60
            tc = self.temp_conductor(_initial_temperature, solve_t, 15, _calc_tau)
            threshold = _conductor_temp_emergency - tc
            _int += 1
            if _int >= _max_iterations:
                _max = True
            if debug:
                print(
                    f'Final Current: {_final_current}, solve_t: {solve_t}, calc_tau: {_calc_tau}, tc: {_tc}, threshold: {threshold}')
                input(f'Iteration #: {_int} Press Enter to continue...')

        if _max:
            print('Unable to converge')
            print(f'Final result: Threshold: {threshold}....Solved input: {solve_t}....Number of Iterations {_int}')
            _results = _final_current
        else:
            _results = _final_current
        if debug:
            print(
                f'Max iterations: {_int}, Final Current: {_final_current}, Threshold: {threshold}, Solved input: {solve_t}')
            print(f'Conductor Temp: {_results}, tau: {_calc_tau}')

        return _results


if __name__ == "__main__":
    app = IEEE738()
    # app.runTest()
    conductor_temp = 100
    amb = 40
    diamm = 28.1
    diai = diamm/25.4
    elevationM = 100
    elevationI = elevationM * 3.281

    print(app.c_uf('metric', conductor_temp, amb))
    print(app.c_pf('metric', conductor_temp, amb, elevationM,))
    print(app.c_kf('metric', conductor_temp, amb))
    print(app.c_qcHeatLoss('metric', diamm, conductor_temp, amb, 100, 90, 0.61))
    print(app.c_qrHeatLoss('metric', diamm, 0.5, conductor_temp, amb))
    print(app.c_qsHeatGain('metric', 0.5, 0.57, 'dirty', 30, 10, 6, 2009, 1100, 'N-S', diamm/1000))

    print(" ")
    print(app.c_uf('imperial', conductor_temp, amb))
    print(app.c_pf('imperial', conductor_temp, amb, elevationI))
    print(app.c_kf('imperial', conductor_temp, amb))
    print(app.c_qcHeatLoss('imperial', diai, conductor_temp, amb, elevationI, 90, 7204.724))
    print(app.c_qrHeatLoss('imperial', diai, 0.5, conductor_temp, amb))
    print(app.c_qsHeatGain('imperial', 0.5, 0.57, 'dirty', 30, 10, 6, 2009, 1100, 'N-S', 1.108 / 12))