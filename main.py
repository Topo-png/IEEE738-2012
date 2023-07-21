# -*- coding: utf-8 -*-
"""
MIT License

Copyright (c) 2023 Mark Shuck

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

@author: Mark Shuck
email: mark@shuck.engineering
"""

import numpy as np
import pandas as pd
import datetime
import time
import scipy.optimize as optimize

import UnitConversion as Unc
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# todo clean up naming conventions

uc = Unc.UnitConvert()

ver = 'v0.2.5'
demo = False  # disable to allow selection of config/conductors

dir_config = 'Sample/config-sample.xlsx'  # location of configuration file
dir_conductor = 'Sample/Conductor_Prop-Sample.xlsx'  # location of conductor file

degree_sign = u'\N{DEGREE SIGN}'


class IEEE738:

    def __init__(self):
        self.true_to_standard = False
        self.true_to_spreadsheet = False  # TODO implement this feature. spreadsheet calculates resistances slightly
        # different compared to the standard
        self.conductor_temp_steps = 6

    def runTest(self):

        t0 = time.time()
        config_list = self.import_config(dir_config)
        conductor_list, conductor_spec_temp_list = self.import_conductor(dir_conductor)

        df_config = self.select_config(config_list)
        df_conductor, df_spec = self.select_conductor(conductor_list, conductor_spec_temp_list)
        df_adjusted = self.unit_conversion(df_conductor, df_spec, df_config)

        df = self.add_calc_columns(df_adjusted)
        df_normal, df_emergency, df_load = self.c_reporting(df)

        df_normal_day = df_normal.pivot(index='conductor temperature', columns='ambient air temperature',
                                        values='rating daytime')
        df_normal_night = df_normal.pivot(index='conductor temperature', columns='ambient air temperature',
                                          values='rating nighttime')
        df_emergency_day = df_emergency.pivot(index='conductor temperature', columns='ambient air temperature',
                                              values='rating daytime')
        df_emergency_night = df_emergency.pivot(index='conductor temperature', columns='ambient air temperature',
                                                values='rating nighttime')
        df_load_day = df_load.pivot(index='conductor temperature', columns='ambient air temperature',
                                    values='load dump rating daytime')
        df_load_night = df_load.pivot(index='conductor temperature', columns='ambient air temperature',
                                      values='load dump rating nighttime')

        print(df_normal_day)
        print(df_normal_night)
        print(df_emergency_day)
        print(df_emergency_night)
        print(df_load_day)
        print(df_load_night)
        self.export_excel(df_normal, df_emergency, df_load, df_config, 'export_test')
        t1 = time.time()
        t = t1 - t0
        print(f'time to compute {t} seconds')
        return None

    @staticmethod
    def import_config(dir_):
        # Imports list of configurations from database
        # Returns list of configurations
        config_list = pd.read_excel(io=dir_, sheet_name='config', engine='openpyxl')
        return config_list

    @staticmethod
    def import_conductor(dir_):
        # Imports list of conductors from database
        # Returns list of conductors and list of conductor specs with Normal/Emergency temperature ratings
        conductor_list = pd.read_excel(io=dir_, sheet_name='conductors', engine='openpyxl')
        conductor_spec = pd.read_excel(io=dir_, sheet_name='conductor spec', engine='openpyxl')
        conductor_list.sort_values('Metal OD', ascending=True, inplace=True)
        conductor_spec.sort_values('Conductor Spec', ascending=True, inplace=True)
        return conductor_list, conductor_spec

    @staticmethod
    def is_number(val):
        try:
            out = float(val)
            return out
        except ValueError:
            return val

    @staticmethod
    def is_number(val):
        try:
            out = float(val)
            return out
        except ValueError:
            return val

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

            if demo:
                _response = 1
            else:
                _response = int(input("Selection? "))
            print(f'{_df_dict[_response - 1]}')
            config = df_config[df_config['config name'] == _df[_response - 1]].reset_index(drop=True)
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
    def select_conductor(df_conductor_list, df_conductor_spec_temp):
        _config_name = None
        _conductor_spec = None
        _conductor_size = None
        _conductor_stranding = None
        _conductor_core_stranding = None
        _response = None

        # Select conductor spec
        _df = df_conductor_list.drop_duplicates(['Conductor Spec'])
        _df = _df['Conductor Spec'].values
        for _pos, _text in enumerate(_df):
            print(f"{_pos + 1}: {_text}")
        if demo:
            _response = 2
        else:
            _response = int(input("Selection? "))
        _conductor_spec = _df[_response - 1]
        print(_conductor_spec)

        # Select conductor size
        _df = df_conductor_list[df_conductor_list['Conductor Spec'] == _conductor_spec].drop_duplicates(['Size'])
        _df = _df['Size'].values
        for _pos, _text in enumerate(_df):
            print(f"{_pos + 1}: {_text}")
        if demo:
            _response = 1
        else:
            _response = int(input("Selection? "))
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
            if demo:
                _response = 4
            else:
                _response = int(input("Selection? "))
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
        df_spec = df_conductor_spec_temp.loc[df_conductor_spec_temp['Conductor Spec'] == _conductor_spec].reset_index(
            drop=True)
        df_conductor = conductor_data.reset_index(drop=True)
        return df_conductor, df_spec

    # todo use names in found standard
    def add_calc_columns(self, df):
        data = ['qc heat loss', 'qc0', 'qc1', 'qc2', 'uf', 'kf', 'pf', 'Qse', 'theta', 'hc: solar altitude', 'delta',
                'omega', 'chi', 'qs heat gain', 'solar altitude correction factor', 'qr heat loss', 'day of year',
                'k angle', 'solar azimuth constant', 'solar azimuth', 'conductor temperature']
        df = pd.concat([df.reset_index(drop=True), pd.DataFrame(columns=data)], axis=1)
        return df

    def unit_conversion(self, df_conductor, df_spec, df_config):

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

        if uc.units_lookup[calculation_units] == uc.metric_value:
            unit_selection = 2
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
            unit_selection = 3

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
            ('high resistance temperature', 'resistance temperature units', 'C', 'C')
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
            value = uc.speed_convert(df_config.at[0, x[0]], df_config.at[0, x[1]], x[unit_selection])
            df_conductor_wind_list.at[0, x[0]] = value
            df_conductor_wind_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_wind_angle_list:
            value = uc.angle_convert(df_config.at[0, x[0]], df_config.at[0, x[1]], x[unit_selection])
            df_conductor_wind_angle_list.at[0, x[0]] = value
            df_conductor_wind_angle_list.at[0, x[1]] = x[unit_selection]
            df_config_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_config_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_length_list:
            value = uc.length_convert(df_conductor.at[0, x[0]], df_conductor.at[0, x[1]], x[unit_selection])
            df_conductor_length_list.at[0, x[0]] = value
            df_conductor_length_list.at[0, x[1]] = x[unit_selection]
            df_conductor_adjusted.drop(columns=x[0], axis=1, inplace=True)
            df_conductor_adjusted.drop(columns=x[1], axis=1, inplace=True)

        for x in conductor_temp_list:
            value = uc.temp_convert(df_conductor.at[0, x[0]], df_conductor.at[0, x[1]], 'C')
            df_conductor_temp_list.at[0, x[0]] = value
            df_conductor_temp_list.at[0, x[1]] = x[unit_selection]
            df_conductor_adjusted.drop(columns=x[0], axis=1, inplace=True)
            try:
                df_conductor_adjusted.drop(columns=x[1], axis=1, inplace=True)
            except:
                None

        for x in conductor_spec_temp_list:
            value = uc.temp_convert(df_spec.at[0, x[0]], df_spec.at[0, x[1]], 'C')
            df_conductor_spec_temp_list.at[0, x[0]] = value
            df_conductor_spec_temp_list.at[0, x[1]] = x[unit_selection]

        for x in config_temp_list:
            value = uc.temp_convert(df_config.at[0, x[0]], df_config.at[0, x[1]], 'C')
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

    def c_steady_state(self, df, calcType=None, _idx=None):
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
        ambient_air_temp = df.at[_idx, 'ambient air temperature']
        conductor_temperature = df.at[_idx, 'conductor temperature']

        # Conductor setup
        diameter = df.at[_idx, 'Metal OD']

        if uc.units_lookup[calculation_units] == uc.metric_value:
            conductor_projection = diameter / 1000
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
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

        if calcType == 'Emergency':
            conductor_wind_speed = df.at[_idx, 'emergency wind speed']
            conductor_wind_angle = df.at[_idx, 'emergency wind angle']
        else:
            conductor_wind_speed = df.at[_idx, 'normal wind speed']
            conductor_wind_angle = df.at[_idx, 'normal wind angle']

        current_rating_day, current_rating_night = self.c_SSRating(calculation_units, diameter, conductor_temperature,
                                                                   ambient_air_temp, elevation, conductor_wind_angle,
                                                                   conductor_wind_speed, emissivity, solar_absorptivity,
                                                                   atmosphere, latitude, day, month, year, hour,
                                                                   conductor_direction, conductor_projection,
                                                                   conductor_resistance, df, _idx)
        return current_rating_day, current_rating_night

    def c_load_dump(self, df, _idx):
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

        # conductor_spec = df.at[_idx, 'Conductor Spec'] # todo is this needed anymore?
        ambient_air_temp = df.at[_idx, 'ambient air temperature']

        conductor_temp_normal = df.at[_idx, 'normal temperature rating']
        conductor_temp_emergency = df.at[_idx, 'emergency temperature rating']
        conductor_wind_emergency = df.at[_idx, 'emergency wind speed']

        # Conductor setup
        diameter = df.at[_idx, 'Metal OD']

        if uc.units_lookup[calculation_units] == uc.metric_value:
            conductor_projection = diameter / 1000
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
            conductor_projection = diameter / 12

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

        conductor_resistance = pd.DataFrame(d)

        wind_angle = df.at[_idx, 'emergency wind angle']

        # _al, _cu, _stl, _alw
        mcp = self.c_mcp(df.at[_idx, 'Al Weight'] / 1000,
                         df.at[_idx, 'Cu Weight'] / 1000,
                         df.at[_idx, 'St Weight'] / 1000,
                         df.at[_idx, 'Alw Weight'] / 1000)

        duration = df.at[_idx, 'duration (minutes)']

        load_dump_day, load_dump_night = self.load_dump(calculation_units, diameter, conductor_temp_normal,
                                                        conductor_temp_emergency,
                                                        ambient_air_temp, elevation, wind_angle,
                                                        conductor_wind_emergency, emissivity, solar_absorptivity,
                                                        atmosphere, latitude, day, month, year, hour,
                                                        conductor_direction, conductor_projection,
                                                        conductor_resistance, mcp, duration)

        if df is not None:
            df.at[_idx, 'load dump rating daytime'] = load_dump_day
            df.at[_idx, 'load dump rating nighttime'] = load_dump_night
            df.at[_idx, 'load dump duration'] = duration

        return load_dump_day, load_dump_night

    def c_temperature_range(self, df_adjusted):
        # TODO add check to verify design summer/winter are captured within ambient temperature range
        #  if not, adjust range to include
        ambient_lower_range = df_adjusted.at[0, 'ambient air temperature lower range']
        ambient_upper_range = df_adjusted.at[0, 'ambient air temperature upper range']
        ambient_temp_inc = df_adjusted.at[0, 'temperature increment']
        conductor_normal_rating = df_adjusted.at[0, 'normal temperature rating']
        conductor_emergency_rating = df_adjusted.at[0, 'emergency temperature rating']
        # TODO adjust conductor ranges to include steps above and below maximums
        # TODO adjust ambient ranges to include steps above and below maximums

        if conductor_normal_rating == conductor_emergency_rating:
            temp_range_conductor = np.arange(conductor_normal_rating - self.conductor_temp_steps * ambient_temp_inc,
                                             conductor_normal_rating + (
                                                     1 + self.conductor_temp_steps) * ambient_temp_inc,
                                             ambient_temp_inc)
        else:
            temp_range_conductor = np.arange(conductor_normal_rating - self.conductor_temp_steps * ambient_temp_inc,
                                             conductor_emergency_rating + (
                                                     1 + self.conductor_temp_steps) * ambient_temp_inc,
                                             ambient_temp_inc)

        temp_range_ambient = np.arange(ambient_lower_range, ambient_upper_range + ambient_temp_inc, ambient_temp_inc)

        return temp_range_ambient, temp_range_conductor

    def c_reporting(self, df_adjusted):
        _idx = 1
        temp_range_ambient, temp_range_conductor = self.c_temperature_range(df_adjusted)
        _temp_list = (
            ('ambient air temperature', 'ambient air temperature units'),
            ('normal temperature rating', 'normal temperature rating units'),
            ('emergency temperature rating', 'emergency temperature rating units')
        )

        df = pd.DataFrame(df_adjusted)

        total_row = temp_range_ambient.size * temp_range_conductor.size + 1
        df_N = pd.concat([df] * total_row, axis=0, ignore_index=True)
        df_E = pd.concat([df] * total_row, axis=0, ignore_index=True)
        df_L = pd.concat([df] * total_row, axis=0, ignore_index=True)

        for i, element_i in enumerate(temp_range_ambient):
            for j, element_j in enumerate(temp_range_conductor):

                df_N.at[_idx, 'ambient air temperature'] = element_i
                df_N.at[_idx, 'conductor temperature'] = element_j
                df_E.at[_idx, 'ambient air temperature'] = element_i
                df_E.at[_idx, 'conductor temperature'] = element_j
                df_L.at[_idx, 'ambient air temperature'] = element_i
                df_L.at[_idx, 'conductor temperature'] = element_j

                _, _ = self.c_steady_state(df_N, 'Normal', _idx)
                _, _ = self.c_steady_state(df_E, 'Emergency', _idx)
                if j == 0:
                    _, _ = self.c_load_dump(df_L, _idx)
                else:
                    df_L.at[_idx, 'load dump rating daytime'] = df_L.at[_idx - 1, 'load dump rating daytime']
                    df_L.at[_idx, 'load dump rating nighttime'] = df_L.at[_idx - 1, 'load dump rating nighttime']
                _idx = _idx + 1

        # TODO add polynomial regression to replace nan with none zero.
        #  ex HD Copper 500 @ Tc = 55C & Amb = 40C I = nan
        #  mention this or highlight cell somehow

        # remove first row (configuration setup)
        # TODO --> likely a better way to make this work
        df_N = df_N.drop(labels=0, axis='index').reset_index(drop=False)
        df_E = df_E.drop(labels=0, axis='index').reset_index(drop=True)
        df_L = df_L.drop(labels=0, axis='index').reset_index(drop=True)

        return df_N, df_E, df_L

    def export_excel(self, df_n, df_e, df_l, df_config, filename_):
        wb = Workbook()
        ws = wb.active
        filename_ = filename_ + '.xlsx'
        ws.title = 'normal'
        for r in dataframe_to_rows(df_n, index=False, header=True):
            ws.append(r)

        ws = wb.create_sheet('emergency')
        for r in dataframe_to_rows(df_e, index=False, header=True):
            ws.append(r)

        ws = wb.create_sheet('load')
        for r in dataframe_to_rows(df_l, index=False, header=True):
            ws.append(r)

        ws = wb.create_sheet('config')
        for r in dataframe_to_rows(df_config, index=False, header=True):
            ws.append(r)

        wb.save(filename_)

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
        if (qr - qs + qc) > 0:
            rating = np.sqrt((qr - qs + qc) / r)
        else:
            rating = 0
        return rating

    def c_uf(self, calculation_units, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates dynamic viscosity of air and returns results (Pa-s or lb/ft-hr)
        :param calculation_units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Dynamic viscosity of air (Pa-s or lb/ft-hr)
        """
        uf = None
        t_film = (conductor_temp + ambient_air_temp) / 2

        if uc.units_lookup[calculation_units] == uc.metric_value:
            uf = (1.458 * 10 ** -6 * (t_film + 273.15) ** 1.5) / (t_film + 383.4)
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
            uf = (0.00353 * (t_film + 273.15) ** 1.5) / (t_film + 383.4)
        if df is not None:
            df.at[_idx, 'uf'] = uf
        return uf

    def c_kf(self, calculation_units, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates thermal conductivity of air and returns results (W/m*C or W/ft*C)
        :param calculation_units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Thermal conductivity of air (W/m*C or W/ft*C)
        """
        kf = None
        t_film = (conductor_temp + ambient_air_temp) / 2

        if uc.units_lookup[calculation_units] == uc.metric_value:
            kf = 2.424 * 10 ** -2 + 7.477 * 10 ** -5 * t_film - 4.407 * 10 ** -9 * t_film ** 2
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
            kf = 0.007388 + 2.279 * 10 ** -5 * t_film - 1.343 * 10 ** -9 * t_film ** 2
        if df is not None:
            df.at[_idx, 'kf'] = kf
        return kf

    def c_pf(self, calculation_units, conductor_temp, ambient_air_temp, elevation, df=None, _idx=0):
        """
        Calculates air density and returns results (kg/m^3 or lb/ft^3)
        :param calculation_units: Units: 'Metric' or 'Imperial'
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param elevation: elevation of conductors
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Air density (kg/m^3 or lb/ft^3
        """
        pf = None
        t_film = (conductor_temp + ambient_air_temp) / 2
        if uc.units_lookup[calculation_units] == uc.metric_value:
            pf = (1.293 - 1.525 * 10 ** -4 * elevation + 6.379 * 10 ** -9 * elevation ** 2) / (
                    1 + 0.00367 * t_film)
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
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

    def c_Qs(self, calculation_units, atmosphere, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates total solar and sky radiated heat flux rate (W/m^2) and returns results
        :param calculation_units: Units: 'Metric' or 'Imperial'
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

        if uc.units_lookup[calculation_units] == uc.metric_value:
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
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
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

    def c_qse(self, calculation_units, elevation, atmosphere, latitude, day, month, year, hour, df=None, _idx=0):
        """
        Calculates elevation corrected total solar and sky radiated heat flux rate (W/m^2) and returns results
        :param calculation_units: Units: 'Metric' or 'Imperial'
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
        ksolar = self.c_ksolar(calculation_units, elevation, df, _idx)
        Qs = self.c_Qs(calculation_units, atmosphere, latitude, day, month, year, hour, df, _idx)
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

    def c_ksolar(self, calculation_units, elevation, df=None, _idx=0):
        """
        Calculates solar heat multiplying factor (kSolar) and returns results
        :param calculation_units: Units: 'Metric' or 'Imperial'
        :param elevation: elevation of conductors
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Solar Heat multiplying factor
        """
        aks = 1.0
        bks = 3.500E-5
        cks = -1.000E-9
        solar_heat_factor = None

        if uc.units_lookup[calculation_units] == uc.metric_value:
            # meters
            if elevation < 1000:
                solar_heat_factor = 1.0
            elif 1000 <= elevation < 2000:
                solar_heat_factor = 1.10
            elif 2000 <= elevation < 4000:
                solar_heat_factor = 1.19
            elif 4000 <= elevation:
                solar_heat_factor = 1.28
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
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
            'South-north': direction_lookup_value_ns,

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

    def c_qsHeatGain(self, calculation_units, solar_absorptivity, elevation, atmosphere, latitude, day, month,
                     year, hour, conductor_direction, conductor_projection, df=None, _idx=0):
        """
        Calculates heat gain rate from the sun and returns results (W/m or W/ft)
        :param calculation_units: Units: 'Metric' or 'Imperial'
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
        qse = self.c_qse(calculation_units, elevation, atmosphere, latitude, day, month, year, hour, df, _idx)
        theta = self.c_Theta(latitude, day, month, year, hour, conductor_direction, df, _idx)
        qs_heat_gain = solar_absorptivity * qse * np.sin(np.radians(theta)) * conductor_projection

        if df is not None:
            df.at[_idx, 'qs heat gain'] = qs_heat_gain
        return qs_heat_gain

    def c_qrHeatLoss(self, calculation_units, diameter, emissivity, conductor_temp, ambient_air_temp, df=None, _idx=0):
        """
        Calculates radiated heat loss rate per unit length and returns results (W/m or W/ft)
        :param calculation_units: Units: 'Metric' or 'Imperial'
        :param diameter: Conductor diameter (mm or in)
        :param emissivity: Emissivity
        :param conductor_temp: Conductor temperature (C)
        :param ambient_air_temp: Ambient air temperature (C)
        :param df: Dataframe holding output conductor/config/calculated values
        :param _idx: index (row) for dataframe
        :return: Radiated heat loss rate per unit length (W/m or W/ft)
        """
        qr = None
        if uc.units_lookup[calculation_units] == uc.metric_value:
            # W/meters
            qr = 0.0178 * diameter * emissivity * (
                    ((conductor_temp + 273.15) / 100) ** 4 - ((ambient_air_temp + 273.15) / 100) ** 4)
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
            # W/feet
            qr = 0.138 * diameter * emissivity * (
                    ((conductor_temp + 273.15) / 100) ** 4 - ((ambient_air_temp + 273.15) / 100) ** 4)

        if df is not None:
            df.at[_idx, 'qr heat loss'] = qr
        return qr

    def c_qcHeatLoss(self, calculation_units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                     wind_speed, df=None, _idx=0):
        """
        Calculates convected heat loss rate per unit length and returns results (W/m or W/ft)
        :param calculation_units: Units: 'Metric' or 'Imperial'
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
        qc0 = None
        qc1 = None
        qc2 = None
        qc_heat_loss = None
        k_angle = self.c_k_angle(wind_angle, df, _idx)
        uf = self.c_uf(calculation_units, conductor_temp, ambient_air_temp, df, _idx)
        kf = self.c_kf(calculation_units, conductor_temp, ambient_air_temp, df, _idx)
        pf = self.c_pf(calculation_units, conductor_temp, ambient_air_temp, elevation, df, _idx)

        if uc.units_lookup[calculation_units] == uc.metric_value:
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
        elif uc.units_lookup[calculation_units] == uc.imperial_value:
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

    def c_SSRating(self, calculation_units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                   wind_speed, emissivity, solar_absorptivity, atmosphere, latitude, day, month, year, hour,
                   conductor_direction, conductor_projection, conductor_resistance, df=None, _idx=0):
        """
        Calculates steady state current and returns results (Amps)
        :param calculation_units: Units: 'Metric' or 'Imperial'
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
        :return: Steady state current (Amps) Day rating includes solar heat gain, Night Rating does not include solar heat gain
        """
        qc = self.c_qcHeatLoss(calculation_units, diameter, conductor_temp, ambient_air_temp, elevation, wind_angle,
                               wind_speed, df, _idx)
        qr = self.c_qrHeatLoss(calculation_units, diameter, emissivity, conductor_temp, ambient_air_temp, df, _idx)
        qs = self.c_qsHeatGain(calculation_units, solar_absorptivity, elevation, atmosphere, latitude, day, month, year,
                               hour, conductor_direction, conductor_projection, df, _idx)
        r_cond = self.c_cond_resistance(conductor_temp, conductor_resistance, df, _idx)
        rating_day = self.current_steady_state(qr, qs, qc, r_cond)
        rating_night = self.current_steady_state(qr, 0, qc, r_cond)

        if df is not None:
            df.at[_idx, 'rating daytime'] = rating_day
            df.at[_idx, 'rating nighttime'] = rating_night
        return rating_day, rating_night

    @staticmethod
    def c_mcp(al, cu, stl, alw):
        results = 433 * al + 192 * cu + 216 * stl + 242 * alw
        return results

    def c_initial_temp(self, calculation_units, diameter, conductor_temp_normal, conductor_temp_emergency,
                       ambient_air_temp, elevation, wind_angle, wind_speed, emissivity, solar_absorptivity,
                       atmosphere, latitude, day_, month_, year_, hour_, conductor_direction, conductor_projection,
                       conductor_resistance):

        initial_current_day, initial_current_night = self.c_SSRating(calculation_units, diameter, conductor_temp_normal,
                                                                     ambient_air_temp, elevation, wind_angle, 0,
                                                                     emissivity, solar_absorptivity, atmosphere,
                                                                     latitude, day_, month_, year_, hour_,
                                                                     conductor_direction, conductor_projection,
                                                                     conductor_resistance)

        result_day = optimize.minimize_scalar(self.c_find_initial_temp,
                                              bounds=(ambient_air_temp, conductor_temp_emergency),
                                              method='bounded',
                                              args=(
                                                  calculation_units, diameter, ambient_air_temp, elevation, wind_angle,
                                                  wind_speed, emissivity, solar_absorptivity, atmosphere, latitude,
                                                  day_, month_, year_, hour_, conductor_direction, conductor_projection,
                                                  conductor_resistance, initial_current_day, 'Day'))

        result_night = optimize.minimize_scalar(self.c_find_initial_temp,
                                                bounds=(ambient_air_temp, conductor_temp_emergency),
                                                method='bounded',
                                                args=(
                                                    calculation_units, diameter, ambient_air_temp, elevation,
                                                    wind_angle, wind_speed, emissivity, solar_absorptivity, atmosphere,
                                                    latitude, day_, month_, year_, hour_, conductor_direction,
                                                    conductor_projection, conductor_resistance, initial_current_night,
                                                    "Night"))

        return result_day.x, result_night.x

    def c_find_initial_temp(self, t_c, calculation_units, diameter, ambient_air_temp, elevation, wind_angle,
                            conductor_wind_emergency, emissivity, solar_absorptivity, atmosphere, latitude, day, month,
                            year,
                            hour, conductor_direction, conductor_projection, conductor_resistance, initial_current,
                            condition):

        current_rating_day, current_rating_night = self.c_SSRating(calculation_units, diameter, t_c, ambient_air_temp,
                                                                   elevation, wind_angle, conductor_wind_emergency,
                                                                   emissivity, solar_absorptivity, atmosphere, latitude,
                                                                   day, month, year, hour, conductor_direction,
                                                                   conductor_projection, conductor_resistance)
        if condition == 'Day':
            delta = initial_current - current_rating_day
        elif condition == 'Night':
            delta = initial_current - current_rating_night

        delta = np.abs(delta)
        return delta

    def c_findTemp(self, initial_temperature, final_temperature, time_, calc_tau):
        results = initial_temperature + (final_temperature - initial_temperature) * (1 - np.exp(-time_ / calc_tau))
        return results

    def find_conductor_temp(self, t_c, calculation_units, diameter, ambient_air_temp, elevation, wind_angle,
                            conductor_wind_emergency, emissivity, solar_absorptivity, atmosphere, latitude, day, month,
                            year, hour, conductor_direction, conductor_projection, conductor_resistance,
                            initial_temperature, initial_current, conductor_temp_emergency, mcp, condition, duration):

        if condition == 'Day':
            final_, _ = self.c_SSRating(calculation_units, diameter, t_c, ambient_air_temp,
                                        elevation, wind_angle, conductor_wind_emergency, emissivity, solar_absorptivity,
                                        atmosphere, latitude, day, month, year, hour, conductor_direction,
                                        conductor_projection, conductor_resistance)
        elif condition == 'Night':
            _, final_ = self.c_SSRating(calculation_units, diameter, t_c, ambient_air_temp,
                                        elevation, wind_angle, conductor_wind_emergency, emissivity, solar_absorptivity,
                                        atmosphere, latitude, day, month, year, hour, conductor_direction,
                                        conductor_projection, conductor_resistance)

        if not self.true_to_standard:
            # TODO fix this to match, add in true to spreadsheet option???
            # todo clean up temperature references
            r = self.c_cond_resistance(initial_temperature, conductor_resistance)
        else:
            r = self.c_cond_resistance(initial_temperature, conductor_resistance)

        calc_tau = (mcp * (t_c - initial_temperature)) / (r * (final_ ** 2 - initial_current ** 2)) * 1 / 60

        tc = (initial_temperature + (t_c - initial_temperature) * (1 - np.exp(-duration / calc_tau)))
        delta = np.abs(conductor_temp_emergency - tc)
        return delta

    def load_dump(self, calculation_units, diameter, conductor_temp_normal, conductor_temp_emergency,
                  ambient_air_temp, elevation, wind_angle, wind_speed, emissivity, solar_absorptivity,
                  atmosphere, latitude, day_, month_, year_, hour_, conductor_direction, conductor_projection,
                  conductor_resistance, mcp, duration):

        initial_current_day, initial_current_night = \
            self.c_SSRating(calculation_units, diameter, conductor_temp_normal,
                            ambient_air_temp, elevation, wind_angle, 0, emissivity, solar_absorptivity, atmosphere,
                            latitude, day_, month_, year_, hour_, conductor_direction, conductor_projection,
                            conductor_resistance)

        initial_temperature_day, initial_temperature_night = \
            self.c_initial_temp(calculation_units, diameter, conductor_temp_normal, conductor_temp_emergency,
                                ambient_air_temp, elevation, wind_angle, wind_speed, emissivity, solar_absorptivity,
                                atmosphere, latitude, day_, month_, year_, hour_, conductor_direction,
                                conductor_projection, conductor_resistance)

        result_day = optimize.minimize_scalar(self.find_conductor_temp, bounds=(ambient_air_temp, 600),
                                              method='bounded',
                                              args=(
                                                  calculation_units, diameter, ambient_air_temp, elevation, wind_angle,
                                                  wind_speed, emissivity, solar_absorptivity, atmosphere, latitude,
                                                  day_, month_, year_, hour_, conductor_direction, conductor_projection,
                                                  conductor_resistance, initial_temperature_day, initial_current_day,
                                                  conductor_temp_emergency, mcp, 'Day', duration))
        result_night = optimize.minimize_scalar(self.find_conductor_temp, bounds=(ambient_air_temp, 600),
                                                method='bounded',
                                                args=(
                                                    calculation_units, diameter, ambient_air_temp, elevation,
                                                    wind_angle, wind_speed, emissivity, solar_absorptivity, atmosphere,
                                                    latitude, day_, month_, year_, hour_, conductor_direction,
                                                    conductor_projection, conductor_resistance,
                                                    initial_temperature_night, initial_current_night,
                                                    conductor_temp_emergency, mcp, 'Night', duration))

        final_temperature_day = result_day.x
        final_temperature_night = result_night.x

        final_current_day, _ = self.c_SSRating(calculation_units, diameter, final_temperature_day, ambient_air_temp,
                                               elevation, wind_angle, wind_speed, emissivity, solar_absorptivity,
                                               atmosphere, latitude, day_, month_, year_, hour_, conductor_direction,
                                               conductor_projection, conductor_resistance)
        _, final_current_night = self.c_SSRating(calculation_units, diameter, final_temperature_night, ambient_air_temp,
                                                 elevation, wind_angle, wind_speed, emissivity, solar_absorptivity,
                                                 atmosphere, latitude, day_, month_, year_, hour_, conductor_direction,
                                                 conductor_projection, conductor_resistance)

        return final_current_day, final_current_night


if __name__ == "__main__":
    app = IEEE738()
    app.runTest()
