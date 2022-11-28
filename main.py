# -*- coding: utf-8 -*-
"""
Created on Sun Apr  3 22:33:02 2022

@author: shuck
"""

import numpy as np
import pandas as pd
import datetime
import time

ver = 'v0.1.0'
debug = False

dir_config = 'config-sample.xlsx' # location of configuration file
dir_conductor = 'Conductor_Prop-Sample.xlsx' # location of conductor file

degree_sign = u'\N{DEGREE SIGN}'


class IEEE738:

    def __init__(self):
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
        config_selection = self.select_config(config_list)
        conductor_selection = self.select_conductor(conductor_list, conductor_spec)
        df = self.add_calculated_values(config_selection, conductor_selection)
        n, e, load, df_out = self.c_reporting(df, 0)
        print(f'Normal ratings {n}')
        print(f'Emergency ratings {e}')
        print(f'load dumping ratings {load}')
        excel_sheets = ('Normal', 'Emergency', 'Load Dump', 'Results')
        self.toExcel(n, e, load, df_out, excel_sheets)
        t1 = time.time()
        t = t1 - t0
        print(f'time to compute {t} seconds')
        return None

    @staticmethod
    def import_config(_dir):
        # Imports list of configurations from database
        # Returns list of configurations
        _df = pd.io.api.ExcelFile(_dir, 'openpyxl')
        config_list = pd.read_excel(_df, sheet_name='config')
        return config_list

    @staticmethod
    def import_conductor(_dir):
        # Imports list of conductors from database
        # Returns list of conductors and list of conductor specs with Normal/Emergency temperature ratings
        _df = pd.io.api.ExcelFile(_dir, 'openpyxl')
        conductor_list = pd.read_excel(_df, sheet_name='conductors')
        conductor_spec = pd.read_excel(_df, sheet_name='conductor spec')
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
            config = df_config[df_config['config name'] == _df[_response - 1]]
            config = df_config[df_config['config name'] == _df[_response - 1]]
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

    # todo make this better, add to all calculated variables
    def add_calculated_values(self, config_selection, conductor_selection):
        df = pd.concat([config_selection.reset_index(drop=True), conductor_selection.reset_index(drop=True)], axis=1)
        data = ['qc', 'qc0', 'qc1', 'qc2', 'uf', 'kf', 'pf', 'He', 'elevation', 'conductor temperature',
                'day of year', 'k_angle']
        df = pd.concat([df, pd.DataFrame(columns=data)])
        return df

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
                _df = _df['Core Strand'].values
                for _pos, _text in enumerate(_df):
                    print(f"{_pos + 1}: {_text}")
                _response = int(input("Selection?"))
                _conductor_core_stranding = _df[_response - 1]
                _df = _df[_df['Core Strand'] == _conductor_core_stranding]

        _df = df_conductor_list  # Reset list to original excel import, could probably do this neater, but it works

        if _conductor_core_stranding is None:
            if _conductor_stranding is None:
                conductor_data = _df.loc[
                    (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size)]
            else:
                conductor_data = _df.loc[
                    (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size) & (
                            _df['Cond Strand'] == _conductor_stranding)]
        elif _conductor_stranding is None:
            conductor_data = _df.loc[
                (_df['Conductor Spec'] == _conductor_spec) & (_df['Size'] == _conductor_size) & (
                        _df['Cond Strand'] == _conductor_stranding) & (_df['Core Strand'] == _conductor_core_stranding)]

        _df_spec = pd.DataFrame()
        _df_spec = df_spec.loc[df_spec['Conductor Spec'] == _conductor_spec]

        conductor_data = pd.concat([conductor_data.reset_index(drop=True),
                                    _df_spec['normal temperature rating'].reset_index(drop=True),
                                    _df_spec['normal temperature rating units'].reset_index(drop=True),
                                    _df_spec['emergency temperature rating'].reset_index(drop=True),
                                    _df_spec['emergency temperature rating units'].reset_index(drop=True)], axis=1)
        return conductor_data

    @staticmethod
    def wind_units_adjustment(_wind, _calculations_units, _config_units):
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

        _metric_value = 0
        _imperial_value = 1

        _wind_speed_lookup_value_mps = 0  # meters per second
        _wind_speed_lookup_value_kmh = 1  # kilometers per hour
        _wind_speed_lookup_value_fps = 2  # feet per second
        _wind_speed_lookup_value_fph = 3  # feet per second
        _wind_speed_lookup_value_mph = 4  # miles per hour
        _wind_speed_lookup_value_knots = 5  # knots

        _units_lookup = {
            'metric': _metric_value,
            'Metric': _metric_value,
            'M': _metric_value,
            'm': _metric_value,
            'Imperial': _imperial_value,
            'imperial': _imperial_value,
            'I': _imperial_value,
            'i': _imperial_value
        }

        _units_output = {
            _metric_value: 'm/s',
            _imperial_value: 'ft/hr'
        }

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
            if _wind <= 0:
                return [0, _units_output[_units_lookup[_calculations_units]]]
            else:
                if _units_lookup[_calculations_units] == _metric_value:  # metric report
                    if _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_mps:
                        return [_wind, _units_output[_units_lookup[_calculations_units]]]  # meters per second
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_kmh:
                        return [_wind / kph_2_mps, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_fps:
                        return [_wind / fps_2_mps, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_fph:
                        return [_wind / fph_2_mps, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_mph:
                        return [_wind / mph_2_mps, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_knots:
                        return [_wind / kn_2_mps, _units_output[_units_lookup[_calculations_units]]]
                elif _units_lookup[_calculations_units] == _imperial_value:  # imperial report
                    if _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_mps:
                        return [_wind / mps_2_fph, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_kmh:
                        return [_wind / kph_2_fph, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_fps:
                        return [_wind / fps_2_fph, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_fph:
                        return [_wind, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_mph:
                        return [_wind / mph_2_fph, _units_output[_units_lookup[_calculations_units]]]
                    elif _wind_speed_lookup[_config_units] == _wind_speed_lookup_value_knots:
                        return [_wind / kn_2_fph, _units_output[_units_lookup[_calculations_units]]]
                else:
                    return [0, _units_output[_units_lookup[_calculations_units]]]
        except KeyError:
            return "Error", "Error"

    def temp_units_adjustment(self, _temp, config_units):
        try:
            if self.temperature_lookup[config_units] == self.temp_lookup_value_C:  # return Celsius
                return [_temp, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_F:  # Fahrenheit to Celsius
                return [(_temp - 32) * 5 / 9, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_K:  # Kelvin to Celsius
                return [_temp + 273.15, self.units_output_temp[0]]
            elif self.temperature_lookup[config_units] == self.temp_lookup_value_R:  # Rankine to Celsius
                return [(_temp - 491.67) * 5 / 9, self.units_output_temp[0]]
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

        # length_lookup_value_mm = 0
        # length_lookup_value_cm = 1
        # length_lookup_value_dm = 2
        # length_lookup_value_m = 3
        # length_lookup_value_mil = 4
        # length_lookup_value_inch = 5
        # length_lookup_value_foot = 6
        # length_lookup_value_mile = 7

        try:
            if self.units_output[unit_type] == self.metric_value:  # metric report
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
            elif self.units_output[unit_type] == self.imperial_value:  # imperial report
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

    @staticmethod
    def diameter_units_adjustment(length, input_units, output_units, unit_type):
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

        metric_value = 0
        imperial_value = 1

        units_lookup = {
            'metric': metric_value,
            'Metric': metric_value,
            'M': metric_value,
            'm': metric_value,
            'Imperial': imperial_value,
            'imperial': imperial_value,
            'I': imperial_value,
            'i': imperial_value
        }

        units_output = {
            metric_value: 'm',
            imperial_value: 'ft'
        }

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
            if units_lookup[unit_type] == metric_value:  # metric report
                if length_lookup[input_units] == length_lookup_value_mm:
                    return [length * mm_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_cm:
                    return [length * cm_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_dm:
                    return [length * dm_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_m:
                    return [length, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mil:
                    return [length * mil_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_inch:
                    return [length * in_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_foot:
                    return [length * foot_2_m, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mile:
                    return [length * mile_2_m, units_output[units_lookup[unit_type]]]
            elif units_lookup[unit_type] == imperial_value:  # imperial report
                if length_lookup[input_units] == length_lookup_value_mm:
                    return [length * mm_2_in, units_output[units_lookup[unit_type]]]  # meters per second
                elif length_lookup[input_units] == length_lookup_value_cm:
                    return [length * cm_2_in, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_dm:
                    return [length * dm_2_in, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_m:
                    return [length * m_2_in, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mil:
                    return [length * mil_2_in, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_inch:
                    return [length, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_foot:
                    return [length * ft_2_in, units_output[units_lookup[unit_type]]]
                elif length_lookup[input_units] == length_lookup_value_mile:
                    return [length * mile_2_in, units_output[units_lookup[unit_type]]]
            else:
                return [0, units_output[units_lookup[unit_type]]]
        except KeyError:
            return "Error", "Error"

    def _unit_conversion(self, df, _idx):
        unit_selection = None
        _df = pd.DataFrame
        df_length = pd.DataFrame()
        df_diameter = pd.DataFrame()
        df_temp = pd.DataFrame()
        df_wind = pd.DataFrame()

        calculation_units = df.at[0, 'calculation units']

        metric_value = 0
        imperial_value = 1

        units_lookup = {
            'metric': metric_value,
            'Metric': metric_value,
            'M': metric_value,
            'm': metric_value,
            'Imperial': imperial_value,
            'imperial': imperial_value,
            'I': imperial_value,
            'i': imperial_value
        }

        if units_lookup[calculation_units] == metric_value:
            unit_selection = 2
        elif units_lookup[calculation_units] == imperial_value:
            unit_selection = 3

        # value, config unit, required unit for metric calculation, required unit for imperial calculation
        _wind_list = (
            ('normal wind speed', 'normal wind speed units', 'deg', 'deg'),
            ('emergency wind speed', 'emergency wind speed units', 'deg', 'deg'),
        )
        _length_list = (
            ('elevation', 'elevation units', 'm', 'ft'),
            ('resistance distance', 'resistance distance units', 'm', 'ft')
        )
        _diameter_list = (
            ('Cond Wire Diameter', 'Cond Wire Diameter Units', 'm', 'in'),
            ('Core Wire Diameter', 'Core Wire Diameter Units', 'm', 'in'),
            ('Core OD', 'Core OD Units', 'm', 'in'),
            ('Metal OD', 'Metal OD Units', 'm', 'in'),
        )
        _temp_list = (
            ('low resistance temperature', 'resistance temperature units'),
            ('high resistance temperature', 'resistance temperature units'),
            ('normal temperature rating', 'normal temperature rating units'),
            ('emergency temperature rating', 'emergency temperature rating units'),
            ('ambient air temperature lower range', 'ambient air temperature units'),
            ('ambient air temperature upper range', 'ambient air temperature units'),
            ('temperature increment', 'ambient air temperature units'),
            ('ambient air temperature', 'ambient air temperature units')
        )

        for x in _wind_list:
            (_speed, _units) = self.wind_units_adjustment(df[x[0]].values[0], calculation_units, df[x[1]].values[0])
            df_wind.at[_idx, 'adjusted ' + x[0]] = _speed
            df_wind.at[_idx, 'adjusted ' + x[1]] = _units

        for x in _temp_list:
            (_temp, _units) = self.temp_units_adjustment(df[x[0]].values[0],
                                                         df[x[1]].values[0])
            df_temp.at[_idx, 'adjusted ' + x[0]] = _temp
            df_temp.at[_idx, 'adjusted ' + x[1]] = _units

        for x in _length_list:
            (_length, _units) = self.length_units_adjustment(df[x[0]].values[0], x[unit_selection],
                                                             calculation_units)
            df_length.at[_idx, 'adjusted ' + x[0]] = _length
            df_length.at[_idx, 'adjusted ' + x[1]] = _units

        for x in _diameter_list:
            (_diameter, _units) = self.diameter_units_adjustment(df[x[0]].values[0], df[x[1]].values[0],
                                                                 x[unit_selection], calculation_units)
            df_diameter.at[_idx, 'adjusted ' + x[0]] = _diameter
            df_diameter.at[_idx, 'adjusted ' + x[1]] = _units

        # TODO this will continuously grow DF each time we call it. Need to adjust this
        df = pd.concat([df.reset_index(drop=True), df_wind.reset_index(drop=True), df_temp.reset_index(drop=True),
                        df_length.reset_index(drop=True), df_diameter.reset_index(drop=True)], axis=1, sort=False)
        df = df.fillna(0)

        return df

    def c_steady_state(self, df, _idx):
        # Configuration setup
        _APrime = None
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
        _ambient_air_temp_adjusted = df.at[_idx, 'ambient air temperature']

        # todo remove this hack, rename conductor temp emergency adjust to something about emergency rating
        _conductor_temp_normal_adjusted = df.at[_idx, 'normal temperature rating']
        _conductor_temp_emergency_adjusted = df.at[_idx, 'emergency temperature rating']
        _conductor_wind_normal_adjusted = df.at[_idx, 'adjusted normal wind speed']
        _conductor_wind_emergency_adjusted = df.at[_idx, 'adjusted emergency wind speed']

        conductor_temperature_adjusted = df.at[_idx, 'conductor temperature']

        # _conductor_temp_normal_adjusted = df.at[_idx, 'adjusted normal temperature rating']
        # _conductor_temp_emergency_adjusted = df.at[_idx, 'adjusted emergency temperature rating']
        # _conductor_wind_normal_adjusted = df.at[_idx, 'adjusted normal wind speed']
        # _conductor_wind_emergency_adjusted = df.at[_idx, 'adjusted emergency wind speed']

        # TODO add in check between conductor & ambient temperature,
        #  if same, end. needs to be posted adjusted temperature. add limit 0.001?
        # TODO add checks for day/month/year here in addition to other day of year method.
        # TODO add note in report that states the normal and emergency temp is the same

        # Conductor setup
        # _diameter = df.at[_idx, 'adjusted Metal OD']
        _diameter = df.at[_idx, 'Metal OD']
        # TODO verify this is correct for the units

        if self.units_lookup[_calculation_units] == self.metric_value:
            _APrime = _diameter / 1000
        elif self.units_lookup[_calculation_units] == self.imperial_value:
            _APrime = _diameter / 12

        # d = {'high resistance Ω/unit': [df['high resistance Ω/unit'].values[0]],  # high resistance
        #      'low resistance Ω/unit': [df['low resistance Ω/unit'].values[0]],  # low resistance
        #      'resistance temperature unit': [df['adjusted resistance temperature units'].values[0]],
        #      # resistance distance unit (C, F, K, R)
        #      'high resistance temperature': [df['adjusted high resistance temperature'].values[0]],
        #      # high resistance temp
        #      'low resistance temperature': [df['adjusted low resistance temperature'].values[0]],  # low resistance temp
        #      'resistance distance': [df['resistance distance'].values[0]],  # resistance distance
        #      'resistance distance unit': [df['adjusted resistance distance units'].values[0]]
        #      # resistance distance unit (mile/meter/etc...)
        #      }

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
        # Normal Rating (wind speed and angle variable)
        _wind_angle = df.at[_idx, 'normal wind angle']

        _outputNormal = self.c_SSRating(_calculation_units, _diameter, conductor_temperature_adjusted,
                                        _ambient_air_temp_adjusted,
                                        _elevation, _wind_angle,
                                        _conductor_wind_normal_adjusted, _emissivity, _solar_absorptivity, _atmosphere,
                                        _latitude, _day, _month,
                                        _year, _hour, _conductor_direction, _APrime, _conductor_resistance, df, _idx)

        # Emergency Rating (wind speed and angle variable)
        _wind_angle_emergency = df.at[_idx, 'emergency wind angle']

        _outputEmergency = self.c_SSRating(_calculation_units, _diameter, conductor_temperature_adjusted,
                                           _ambient_air_temp_adjusted, _elevation,
                                           _wind_angle_emergency, _conductor_wind_emergency_adjusted, _emissivity,
                                           _solar_absorptivity,
                                           _atmosphere, _latitude,
                                           _day, _month, _year, _hour, _conductor_direction, _APrime,
                                           _conductor_resistance)

        return _outputNormal, _outputEmergency

    def c_load_dump(self, df, _idx):
        # Configuration setup
        _APrime = None
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
        _ambient_air_temp_adjusted = df.at[_idx, 'ambient air temperature']
        # _conductor_temp_normal_adjusted = df.at[_idx, 'adjusted normal temperature rating']
        # _conductor_temp_emergency_adjusted = df.at[_idx, 'adjusted emergency temperature rating']
        # _conductor_wind_emergency_adjusted = df.at[_idx, 'adjusted emergency wind speed']

        _conductor_temp_normal_adjusted = df.at[_idx, 'normal temperature rating']
        _conductor_temp_emergency_adjusted = df.at[_idx, 'emergency temperature rating']
        _conductor_wind_emergency_adjusted = df.at[_idx, 'emergency wind speed']
        conductor_temperature_adjusted = df.at[_idx, 'conductor temperature']

        # TODO add in check between conductor & ambient temperature,
        #  if same, end. needs to be posted adjusted temperature. add limit 0.001?
        # TODO add checks for day/month/year here in addition to other day of year method.
        # TODO add note in report that states the normal and emergency temp is the same

        # Conductor setup
        # _diameter = df.at[_idx, 'adjusted Metal OD']
        _diameter = df.at[_idx, 'Metal OD']
        # TODO verify this is correct for the units

        if self.units_lookup[_calculation_units] == self.metric_value:
            _APrime = _diameter / 1000
        elif self.units_lookup[_calculation_units] == self.imperial_value:
            _APrime = _diameter / 12

        # d = {'high resistance Ω/unit': [df['high resistance Ω/unit'].values[0]],  # high resistance
        #      'low resistance Ω/unit': [df['low resistance Ω/unit'].values[0]],  # low resistance
        #      'resistance temperature unit': [df['adjusted resistance temperature units'].values[0]],
        #      # resistance distance unit (C, F, K, R)
        #      'high resistance temperature': [df['adjusted high resistance temperature'].values[0]],
        #      # high resistance temp
        #      'low resistance temperature': [df['adjusted low resistance temperature'].values[0]],  # low resistance temp
        #      'resistance distance': [df['resistance distance'].values[0]],  # resistance distance
        #      'resistance distance unit': [df['adjusted resistance distance units'].values[0]]
        #      # resistance distance unit (mile/meter/etc...)
        #      }
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

        load_dump = self.test(_threshold, _calculation_units, _diameter, _conductor_temp_normal_adjusted, _conductor_temp_emergency_adjusted, conductor_temperature_adjusted,
                              _ambient_air_temp_adjusted,
                              _elevation,
                              _wind_angle, _conductor_wind_emergency_adjusted, _emissivity, _solar_absorptivity,
                              _atmosphere,
                              _latitude,
                              _day, _month, _year, _hour, _conductor_direction, _APrime,
                              _conductor_resistance, _tau, _mcp)
        return load_dump

    def c_reporting(self, df, _idx):

        ambient_lower_range = df.at[0, 'ambient air temperature lower range']
        ambient_upper_range = df.at[0, 'ambient air temperature upper range']
        ambient_temp_inc = df.at[0, 'temperature increment']
        conductor_normal_rating = df.at[0, 'normal temperature rating']
        conductor_emergency_rating = df.at[0, 'emergency temperature rating']
        conductor_temp_steps = 6

        if conductor_normal_rating == conductor_emergency_rating:
            temp_range_conductor = np.arange(conductor_normal_rating - conductor_temp_steps * ambient_temp_inc,
                                             conductor_normal_rating + ambient_temp_inc, ambient_temp_inc)
        else:
            temp_range_conductor = np.arange(conductor_normal_rating - conductor_temp_steps * ambient_temp_inc,
                                             conductor_emergency_rating + ambient_temp_inc, ambient_temp_inc)

        temp_range_ambient = np.arange(ambient_lower_range, ambient_upper_range + ambient_temp_inc, ambient_temp_inc)

        _temp_list = (
            ('ambient air temperature', 'ambient air temperature units'),
            ('normal temperature rating', 'normal temperature rating units'),
            ('emergency temperature rating', 'emergency temperature rating units')
        )

        # todo add in emergency conductor rating, right now only normal conductor temperature is being adjusted

        df = self._unit_conversion(df, _idx)
        normal_report = np.empty((temp_range_conductor.shape[0], 0))
        emergency_report = np.empty((temp_range_conductor.shape[0], 0))
        load_report = np.empty((temp_range_conductor.shape[0], 0))
        _idx = _idx + 1
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
                # for x in _temp_list:
                #     _temp, _units = self.temp_units_adjustment(df.at[_idx, x[0]],
                #                                                df.at[_idx, x[1]])
                #     df.at[_idx, 'adjusted ' + x[0]] = _temp
                #     df.at[_idx, 'adjusted ' + x[1]] = _units
                normal_rating, emergency_rating = self.c_steady_state(df, _idx)
                if j == 0:  # todo make this better
                    load_rating = self.c_load_dump(df, _idx)
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

        return df_normal, df_emergency, df_load, df

    def toExcel(self, df_n, df_e, df_l, df_otuput, sheetnames):
        # todo add some note to the DF that shows which calculation it is day/night normal/emergency/load dump
        with pd.ExcelWriter("test.xlsx") as writer:
            df_n.to_excel(writer, sheet_name=sheetnames[0])
            df_e.to_excel(writer, sheet_name=sheetnames[1])
            df_l.to_excel(writer, sheet_name=sheetnames[2])
            df_otuput.to_excel(writer, sheet_name=sheetnames[3])
        return None

    # Functions
    @staticmethod
    def current_steady_state(_qr, _qs, _qc, _r):
        _results = np.sqrt((_qr - _qs + _qc) / _r)
        return _results

    @staticmethod
    def c_uf(_conductor_temp, _ambient_air_temp, df=None, _idx=0):
        # Temperature in degrees Celsius
        if df is None:
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _results = (0.00353 * (_tmp_avg + 273.15) ** 1.5) / (_tmp_avg + 383.4)
        else:
            _conductor_temp = df.at[_idx, 'conductor temperature']
            _ambient_air_temp = df.at[_idx, 'ambient air temperature']
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _results = (0.00353 * (_tmp_avg + 273.15) ** 1.5) / (_tmp_avg + 383.4)
            df.at[_idx, 'uf'] = _results
        return _results

    @staticmethod
    def c_kf(_conductor_temp, _ambient_air_temp, df=None, _idx=0):

        if df is None:
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _results = 0.007388 + (2.279 * 10 ** -5 * _tmp_avg) - 1.343 * 10 ** -9 * _tmp_avg ** 2
        else:
            _conductor_temp = df.at[_idx, 'conductor temperature']
            _ambient_air_temp = df.at[_idx, 'ambient air temperature']
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _results = 0.007388 + (2.279 * 10 ** -5 * _tmp_avg) - 1.343 * 10 ** -9 * _tmp_avg ** 2
            df.at[_idx, 'kf'] = _results

        return _results

    def c_pf(self, _conductor_temp, _ambient_air_temp, _elevation, _units, df=None, _idx=0):

        if df is None:
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _He = self.c_He(_units, _elevation)
            _results = (0.080695 - (2.901 * 10 ** -6) * _He + (3.7 * 10 ** -11) * (_He ** 2)) / (1 + 0.00367 * _tmp_avg)
        else:
            _conductor_temp = df.at[_idx, 'conductor temperature']
            _ambient_air_temp = df.at[_idx, 'ambient air temperature']
            _tmp_avg = (_conductor_temp + _ambient_air_temp) / 2
            _He = self.c_He(_units, _elevation, df, _idx)
            _results = (0.080695 - (2.901 * 10 ** -6) * _He + (3.7 * 10 ** -11) * (_He ** 2)) / (1 + 0.00367 * _tmp_avg)
            df.at[_idx, 'pf'] = _results

        return _results

    @staticmethod
    def c_cond_resistance(_conductor_temp, _conductor_resistance):

        _high_resistance_Ohm_per_unit_distance = _conductor_resistance.at[0, 'high resistance Ω/unit']

        _low_resistance_Ohm_per_unit_distance = _conductor_resistance.at[0, 'low resistance Ω/unit']

        _high_resistance_temperature = _conductor_resistance.at[0, 'high resistance temperature']
        _low_resistance_temperature = _conductor_resistance.at[0, 'low resistance temperature']
        _resistance_temperature_unit = _conductor_resistance.at[0, 'resistance temperature unit']
        _resistance_distance = _conductor_resistance.at[0, 'resistance distance']

        _high_resistance = _high_resistance_Ohm_per_unit_distance / _resistance_distance
        _low_resistance = _low_resistance_Ohm_per_unit_distance / _resistance_distance

        # todo add mention in report that conductor temperature is higher than resistance temperature and the results
        #  are less conservative
        #  overall resistance at temperature might be lower than what actually occurs physically
        #  reference page 10 738-2006
        _r = (
                ((_high_resistance - _low_resistance) /
                 (_high_resistance_temperature - _low_resistance_temperature)) *
                (_conductor_temp - _low_resistance_temperature) + _low_resistance
        )

        return _r

    def c_Qs(self, _atmosphere, _latitude, _day, _month, _year, _hour, df=None, _idx=0):
        if _atmosphere == 'clear':
            # Clear atmosphere
            _A = -3.9241
            _B = 5.9276
            _C = -1.7856E-1
            _D = 3.223E-3
            _E = -3.3549E-5
            _F = 1.8053E-7
            _G = -3.7868E-10
        else:
            # Industrial atmosphere
            _A = 4.9408
            _B = 1.3202
            _C = 6.1444E-2
            _D = -2.9411E-3
            _E = 5.07752E-5
            _F = -4.03627E-7
            _G = 1.22967E-9

        _Hc = self.c_Hc(_latitude, _day, _month, _year, _hour, df, _idx)
        _results = _A + _B * _Hc + _C * _Hc ** 2 + _D * _Hc ** 3 + _E * _Hc ** 4 + _F * _Hc ** 5 + _G * _Hc ** 6

        if debug:
            print(f'Qs (deg): {_results}')
        return _results

    def c_Qse(self, _units, _elevation, _atmosphere, _latitude, _day, _month, _year, _hour, df=None, _idx=0):
        _He = self.c_He(_units, _elevation)
        _kSolar = self.c_kSolar(_He)
        _Qs = self.c_Qs(_atmosphere, _latitude, _day, _month, _year, _hour, df, _idx)
        _results = _kSolar * _Qs
        if debug:
            print(f'Qse (deg): {_results}')
        return _results

    @staticmethod
    def c_day_of_year(_day, _month, _year, df, _idx):
        # TODO add check to verify day is actually in the month & year (leap)
        if df is not None:
            if df.at[_idx, 'day of year'] is not None:
                date = datetime.datetime(int(_year), int(_month), int(_day))
                _results = int(date.strftime("%j"))  # Get the day of the year
                df.at[_idx, 'day of year'] = _results
            _results = df.at[_idx, 'day of year']
        else:
            date = datetime.datetime(int(_year), int(_month), int(_day))
            _results = int(date.strftime("%j"))  # Get the day of the year
        if debug:
            print("Day of year: ", _results)
        return _results, df

    @staticmethod
    def c_kSolar(_he):
        _Aks = 1.0
        _Bks = 3.500E-5
        _Cks = -1.000E-9

        _results = _Aks + _Bks * _he + _Cks * _he * _he
        if debug:
            print("kSolar:", _results)
        return _results

    @staticmethod
    def c_He(_units, _elevation, df=None, _idx=0):
        _results = None
        if _units == 'metric':
            # meters
            if _elevation < 1000:
                _results = 1.0
            elif 1000 <= _elevation < 2000:  #
                _results = 1.10
            elif 2000 <= _elevation < 4000:
                _results = 1.19
            elif 4000 <= _elevation:
                _results = 1.28
        else:
            # feet
            if _elevation < 5000:
                _results = 1.0
            elif 5000 <= _elevation < 10000:
                _results = 1.15
            elif 10000 <= _elevation < 15000:
                _results = 1.25
            elif 15000 <= _elevation:
                _results = 1.30

        if df is not None:
            df.at[_idx, 'He'] = _results

        if debug:
            print("He:", _results)
        return _results

    @staticmethod
    def c_chi(_omega, _latitude, _delta):
        # everything in radians
        _results = np.sin(_omega) / (np.sin(_latitude) * np.cos(_omega) - np.cos(_latitude) * np.tan(_delta))
        if debug:
            print("Chi (rads):", _results)
        return _results

    @staticmethod
    def c_delta(_N):
        # everything in radians
        _results = np.radians(23.4583 * np.sin(np.radians((284 + _N) / 365 * 360)))
        if debug:
            print("Delta (rads):", _results)
        return _results

    @staticmethod
    def c_omega(_hour):
        # everything in radians
        _results = np.radians((_hour / 100 - 12) * 15)
        if debug:
            print("Omega (rads):", _results)
        return _results

    @staticmethod
    def c_CSolar(_omega, _chi):
        _omega = np.degrees(_omega)
        if -180 <= _omega < 0:
            if _chi >= 0:
                _results = 0
            else:
                _results = 180
        else:
            if _chi >= 0:
                _results = 180
            else:
                _results = 360
        if debug:
            print("cSolar (deg):", _results)
        return _results

    def c_Zc(self, _latitude, _day, _month, _year, _hour, df=None, _idx=0):
        # Solar azimuth
        # build out better way of selecting day/month
        # include comment about year and leap years?
        _latitude = np.radians(_latitude)
        _N, _ = self.c_day_of_year(_day, _month, _year, df, _idx)
        _delta = self.c_delta(_N)
        _omega = self.c_omega(_hour)
        _chi = self.c_chi(_omega, _latitude, _delta)
        _CSolar = self.c_CSolar(_omega, _chi)
        _results = _CSolar + np.degrees(np.arctan(_chi))
        if debug:
            print("Zc (deg)", _results)
        return _results

    def c_Hc(self, _latitude, _day, _month, _year, _hour, df=None, _idx=0):

        _latitude = np.radians(_latitude)
        _N, _ = self.c_day_of_year(_day, _month, _year, df, _idx)
        _delta = self.c_delta(_N)
        _omega = self.c_omega(_hour)
        _results = np.degrees(np.arcsin(np.cos(_latitude) * np.cos(_delta) * np.cos(_omega) +
                                        np.sin(_latitude) * np.sin(_delta)))

        if debug:
            print(f'Hc (deg): {_results}, Latitude: {_latitude}, Day: {_day}, Month: {_month}, '
                  f'Year: {_year}, Hour: {_hour}')
        return _results

    def c_Theta(self, _latitude, _day, _month, _year, _hour, _conductor_direction, df=None, _idx=0):
        # todo make this better N/S north/south north-south
        if _conductor_direction == "N/S":
            _Z1 = 0
        else:
            _Z1 = 90

        _Hc = self.c_Hc(_latitude, _day, _month, _year, _hour)
        _Zc = self.c_Zc(_latitude, _day, _month, _year, _hour, df, _idx)

        _results = np.degrees(np.arccos(np.cos(np.radians(_Hc)) * np.cos(np.radians(_Zc - _Z1))))

        if debug:
            print(f'Theta (deg): {_results}, Hc: {_Hc}, Zc: {_Zc}, Latitude: {_latitude}, Day: {_day}, Month: {_month},'
                  f' Year: {_year}, Hour: {_hour}, Conductor Direction: {_conductor_direction}')
        return _results

    @staticmethod
    def c_k_angle(_wind_angle, df=None, _idx=0):
        # Angle between the wind direction and the conductor
        _results = None
        if df is not None:
            if df.at[_idx, 'k_angle'] is not None:
                _results = 1.194 - np.cos(np.radians(_wind_angle)) + \
                           0.194 * np.cos(np.radians(2 * _wind_angle)) + 0.368 * np.sin(np.radians(2 * _wind_angle))
                df.at[_idx, 'k_angle'] = _results
        else:
            _results = 1.194 - np.cos(np.radians(_wind_angle)) + \
                       0.194 * np.cos(np.radians(2 * _wind_angle)) + 0.368 * np.sin(np.radians(2 * _wind_angle))

        # Angle between wind direction and a perpendicular line to the conductor

        if debug:
            print(f'k_angle (deg): {_results}, Wind Angle (deg): {_wind_angle}')
        return _results

    def c_qsHeatGain(self, _solar_absorptivity, _units, _elevation, _atmosphere, _latitude, _day, _month, _year, _hour,
                     _conductor_direction, _aprime, df=None, _idx=0):
        _Qse = self.c_Qse(_units, _elevation, _atmosphere, _latitude, _day, _month, _year, _hour, df, _idx)
        _Theta = self.c_Theta(_latitude, _day, _month, _year, _hour, _conductor_direction, df, _idx)
        _results = _solar_absorptivity * _Qse * np.sin(np.radians(_Theta)) * _aprime
        # TODO verify all of this restores proper values
        if debug:
            print(f'qs Heat Gain: {_results}, Solar Absorptivity: {_solar_absorptivity}, Units: {_units}, '
                  f'Elevation: {_elevation}, Atmosphere: {_atmosphere}, Latitude: {_latitude}, Day: {_day}, '
                  f'Month: {_month}, Year: {_year}, Hour: {_hour}, Conductor direction: {_conductor_direction},'
                  f'Conductor projection: {_aprime}')
        return _results

    @staticmethod
    def c_qrHeatLoss(_units, _diameter, _emissivity, _conductor_temp, _ambient_air_temp, df=None, _idx=0):
        if _units == 'metric':
            # W/meters
            _results = 0.0178 * _diameter * _emissivity * (
                    ((_conductor_temp + 273.15) / 100) ** 4 - ((_ambient_air_temp + 273.15) / 100) ** 4)
        else:
            # W/feet
            _results = 0.138 * _diameter * _emissivity * (
                    ((_conductor_temp + 273.15) / 100) ** 4 - ((_ambient_air_temp + 273.15) / 100) ** 4)
        if debug:
            print(f'qr Heat Loss: {_results}, Units: {_units}, Diameter: {_diameter}, Emissivity: {_emissivity}, '
                  f'Conductor_temp: {_conductor_temp}, Ambient air temp: {_ambient_air_temp}')
        return _results

    def c_qcHeatLoss(self, _units, _diameter, _conductor_temp, _ambient_air_temp, _elevation, _wind_angle,
                     _Vw, df=None, _idx=0):
        k_angle = self.c_k_angle(_wind_angle, df, _idx)
        uf = self.c_uf(_conductor_temp, _ambient_air_temp, df, _idx)
        kf = self.c_kf(_conductor_temp, _ambient_air_temp, df, _idx)
        pf = self.c_pf(_conductor_temp, _ambient_air_temp, _elevation, _units, df, _idx)

        if df is not None:
            uf = df.at[_idx, 'uf']

        if _units == 'metric':
            # W/meter
            # natural convection
            _results_1 = 0.0205 * pf ** 0.5 * _diameter ** 0.75 * (_conductor_temp - _ambient_air_temp) ** 1.25

            # low wind speeds
            _results_2 = (1.01 + 0.0372 * ((_diameter * pf * _Vw) / uf) ** 0.52) * kf * k_angle * (
                    _conductor_temp - _ambient_air_temp)

            # high wind speeds
            _results_3 = (0.0119 * ((_diameter * pf * _Vw) / uf) ** 0.6) * kf * k_angle * (
                    _conductor_temp - _ambient_air_temp)
            _results = np.amax((_results_1, _results_2, _results_3))
        else:
            # _Vw = _Vw * 3600
            # W/feet

            # natural convection
            _results_1 = 0.283 * pf ** 0.5 * _diameter ** 0.75 * (_conductor_temp - _ambient_air_temp) ** 1.25

            # low wind speeds
            _results_2 = (1.01 + 0.371 * ((_diameter * pf * _Vw) / uf) ** 0.52) * kf * k_angle * (
                    _conductor_temp - _ambient_air_temp)

            # high wind speeds
            _results_3 = (0.1695 * ((_diameter * pf * _Vw) / uf) ** 0.6) * kf * k_angle * (
                    _conductor_temp - _ambient_air_temp)
            _results = np.amax((_results_1, _results_2, _results_3))

        if debug:
            print(
                f'dia: {_diameter} pf: {pf} wind: {_Vw} uf: {uf} kf: {kf} wAngle: {k_angle} '
                f'conductorT: {_conductor_temp} ambient: {_ambient_air_temp}')
            print(f'Natural convection:{_results_1} low wind:{_results_2} high wind:{_results_3}')
            print(f'qc Heat Loss (max): {_results}')
        return _results

    def c_SSRating(self, _units, _diameter, _conductor_temp, _ambient_air_temp, _elevation, _wind_angle, _Vw,
                   _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                   _conductor_direction, _APrime, _conductor_resistance, df=None, _idx=0):
        if df is not None:
            df.at[_idx, 'conductor temperature'] = _conductor_temp

        qc = self.c_qcHeatLoss(_units, _diameter, _conductor_temp, _ambient_air_temp, _elevation, _wind_angle, _Vw, df,
                               _idx)
        qr = self.c_qrHeatLoss(_units, _diameter, _emissivity, _conductor_temp, _ambient_air_temp, df, _idx)
        qs = self.c_qsHeatGain(_solar_absorptivity, _units, _elevation, _atmosphere, _latitude, _day, _month, _year,
                               _hour, _conductor_direction, _APrime, df, _idx)
        r_cond = self.c_cond_resistance(_conductor_temp, _conductor_resistance)
        results = self.current_steady_state(qr, qs, qc, r_cond)
        if debug:
            print(f'qc: {qc} qr: {qr} qs: {qs} R: {r_cond}, results: {results}, tc: {_conductor_temp} amb: {_ambient_air_temp}')
            print("Steady state current rating", results)
        return results

    def c_cPF(self, _units, _elevation, _ta, _pf):
        _He = self.c_He(_units, _elevation)
        _results = (2 * (0.080695 - 2.901E-6 * _He + 3.87E-11 * _He ** 2) - 0.00367 * _ta * _pf - 2 * _pf) / (
                0.00367 * _pf)

        if debug:
            print("Conductor Temp", _results)
        return _results

    @staticmethod
    def c_mcp(_al, _cu, _stl, _alw):
        _results = 433 * _al + 192 * _cu + 216 * _stl + 242 * _alw
        return _results

    def iTemp(self, _threshold, _units, _diameter, _conductor_temp_normal, _conductor_temp_emergency, _conductor_temp, _ambient_air_temp, _elevation, _wind_angle, _Vw,
              _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
              _conductor_direction, _APrime, _conductor_resistance):
        _max_iterations = 50
        _max = False
        _int = 0
        _ii = self.c_SSRating(_units, _diameter, _conductor_temp_normal, _ambient_air_temp, _elevation, _wind_angle, 0,
                              _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                              _conductor_direction, _APrime,
                              _conductor_resistance)
        # todo add something here to change wind speed from 0 to none zero but greater than emergency wind speed rating

        lower_t = _ambient_air_temp  # conductor cannot be lower than ambient unless actively cooled
        upper_t = 300  # todo add in some check that this is high enough???
        solve_t = (lower_t + upper_t) / 2
        threshold = _ii - self.c_SSRating(_units, _diameter, solve_t, _ambient_air_temp, _elevation, _wind_angle, _Vw,
                                          _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month,
                                          _year, _hour, _conductor_direction, _APrime, _conductor_resistance)
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
                                     _hour, _conductor_direction, _APrime, _conductor_resistance)
            threshold = _ii - holder
            _int += 1
            if _int >= _max_iterations:
                _max = True
            if debug:
                print(f'Holder: {holder} Threshold: {threshold}')
                input(f'Iteration #: {_int} Press Enter to continue...')

        if _max:
            # todo make mention somewhere that the solver was unable to converge
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
        results = (60 * _tau * _r * (_final_current ** 2 - _initial_current**2)) / _mcp + _initial_temperature
        return results

    def temp_conductor(self, _initial_temperature, _final_temperature, _time, _calc_tau):
        results = _initial_temperature + (_final_temperature - _initial_temperature) * (1 - np.exp(-_time / _calc_tau))
        return results

    def test(self, _threshold, _calculation_units, _diameter, _conductor_temp_normal, _conductor_temp_emergency, _conductor_temp, _ambient_air_temp, _elevation,
             _wind_angle, _conductor_wind_emergency_adjusted,
             _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
             _conductor_direction, _APrime, _conductor_resistance, _tau, _mcp):

        holder = None

        # todo make sure all winds are correct.
        # Initial current, no wind
        # initial temperature, --> initial current with emergency wind applied
        # verify wind angles and naming conventions throughout mix of Vw and normal/emergency

        _max_iterations = 50
        _max = False
        _int = 0
        #todo fix wind unit conversions everywhere
        _initial_temperature = self.iTemp(_threshold, _calculation_units, _diameter, _conductor_temp_normal, _conductor_temp_emergency, _conductor_temp,
                                          _ambient_air_temp,
                                          _elevation,
                                          _wind_angle, _conductor_wind_emergency_adjusted * 3600, _emissivity,
                                          _solar_absorptivity, _atmosphere,
                                          _latitude,
                                          _day, _month, _year, _hour, _conductor_direction, _APrime,
                                          _conductor_resistance)

        # _final_temperature = _initial_temperature / 0.632
        _final_temperature = _conductor_temp_emergency / 0.632

        # calculate initial current

        _initial_current = self.c_SSRating(_calculation_units, _diameter, _conductor_temp_normal, _ambient_air_temp,
                                           _elevation,
                                           _wind_angle, 0,
                                           _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month,
                                           _year, _hour,
                                           _conductor_direction, _APrime, _conductor_resistance)

        _final_current = self.c_SSRating(_calculation_units, _diameter, _final_temperature, _ambient_air_temp,
                                         _elevation,
                                         _wind_angle, _conductor_wind_emergency_adjusted * 3600, _emissivity,
                                         _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year, _hour,
                                         _conductor_direction, _APrime, _conductor_resistance)

        # _r = self.c_cond_resistance((_initial_temperature+_final_temperature)/2, _conductor_resistance)
        # todo pick a method, average/low/Ti
        _r = self.c_cond_resistance(_initial_temperature, _conductor_resistance)
        _calc_tau = (_mcp * (_final_temperature - _initial_temperature)) / (
                _r * (_final_current ** 2 - _initial_current ** 2)) / 60

        _tc = self.temp_conductor(_initial_temperature, _final_temperature, 15, _calc_tau)

        lower_t = _initial_temperature
        # upper_t = _conductor_temp_emergency * 2
        upper_t = 250  # todo pass either highest temperature value from conductor spec table or
        # multiple 2x emergency temp
        solve_t = (lower_t + upper_t) / 2

        threshold = _conductor_temp_emergency - _tc
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
            _r = self.c_cond_resistance(solve_t, _conductor_resistance)
            _final_current = self.c_SSRating(_calculation_units, _diameter, solve_t, _ambient_air_temp, _elevation, _wind_angle,
                                     _conductor_wind_emergency_adjusted*3600,
                                     _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year,
                                     _hour, _conductor_direction, _APrime, _conductor_resistance)
            _calc_tau = (_mcp * (solve_t - _initial_temperature)) / (
                    _r * (_final_current ** 2 - _initial_current ** 2)) / 60
            tc = self.temp_conductor(_initial_temperature, solve_t, 15, _calc_tau)
            threshold = _conductor_temp_emergency - tc
        results = _final_current

        # todo remove everything below if above works properly.....
        # lower_t = _initial_temperature
        # # upper_t = _conductor_temp_emergency * 2
        # upper_t = 250  # todo pass either highest temperature value from conductor spec table or
        # # multiple 2x emergency temp
        # solve_t = (lower_t + upper_t) / 2
        # threshold = _conductor_temp_emergency - _tc #todo fix this
        #
        # if debug:
        #     print(f'Threshold: {threshold}')
        #
        # while np.abs(threshold) >= np.abs(_threshold) and not _max:
        #     if debug:
        #         print(f'range is: {lower_t}  ----  {solve_t}   ----   {upper_t}')
        #     if threshold < 0:
        #         upper_t = solve_t
        #         solve_t = (lower_t + upper_t) / 2
        #         if debug:
        #             print(f'< {solve_t}')
        #     elif threshold > 0:
        #         lower_t = solve_t
        #         solve_t = (lower_t + upper_t) / 2
        #         if debug:
        #             print(f'> {solve_t}')
        #
        #     holder = self.c_SSRating(_calculation_units, _diameter, solve_t, _ambient_air_temp, _elevation, _wind_angle,
        #                              _conductor_wind_emergency_adjusted,
        #                              _emissivity, _solar_absorptivity, _atmosphere, _latitude, _day, _month, _year,
        #                              _hour, _conductor_direction, _APrime, _conductor_resistance)
        #
        #     # _r = self.c_cond_resistance(solve_t, _conductor_resistance)
        #     # todo a few options here, calculate R based on Ti and leave fixed, provides similar number to excel
        #     #  sheet, calculate based on conductor maximum temperature, calculate each time based on iterative
        #     #  approach, calculate based on average (from standard)
        #     _calc_tau = (_mcp * (solve_t - _initial_temperature)) / (
        #             _r * (holder ** 2 - _initial_current ** 2)) / 60
        #     _tc = _initial_temperature + (_conductor_temp_emergency - _initial_temperature) * (1 - np.exp(-15 / _calc_tau))
        #     threshold = _conductor_temp_emergency - _tc #todo fix this
        #     if debug:
        #         print(f'holder: {holder}, solve_t: {solve_t}, calc_tau: {_calc_tau}, tc: {_tc}, threshold: {threshold}')
        #     _int += 1
        #     if _int >= _max_iterations:
        #         _max = True
        #     if debug:
        #         print(f'Holder: {holder} Threshold: {threshold}, Calc Tau: {_calc_tau}, R: {_r}')
        #         input(f'Iteration #: {_int} Press Enter to continue...')
        # if debug:
        #     # todo make mention somewhere that the solver was unable to converge
        #     print(f'Final result: Threshold: {threshold}....Solved input: {solve_t}....Number of Iterations {_int}')
        #     _results = holder
        # else:
        #     _results = holder
        # if debug:
        #     print(f'Max iterations: {_int}, Final result: Threshold: {threshold}....Solved input: {solve_t}')
        #     print(f'Conductor Temp: {_results}, tau: {_calc_tau}')
        return results


if __name__ == "__main__":
    app = IEEE738()
    app.runTest()
