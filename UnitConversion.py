# -*- coding: utf-8 -*-
"""
MIT License

Copyright (c) 2022 Topo-png

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
"""

degree_sign = u'\N{DEGREE SIGN}'


class UnitConvert:

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

        self.temp_lookup_value_C = 0
        self.temp_lookup_value_F = 1
        self.temp_lookup_value_K = 2
        self.temp_lookup_value_R = 3

        self.dict_temp_convert = {
            self.temp_lookup_value_C: {
                self.temp_lookup_value_C: (0, 1, 0),
                self.temp_lookup_value_F: (0, 9/5, 32),
                self.temp_lookup_value_K: (0, 1, 273.15),
                self.temp_lookup_value_R: (0, 9/5, 491.67)
            },
            self.temp_lookup_value_F: {
                self.temp_lookup_value_C: (-32, 5/9, 0),
                self.temp_lookup_value_F: (0, 1, 0),
                self.temp_lookup_value_K: (-32, 5/9, 273.15),
                self.temp_lookup_value_R: (0, 1, 459.67)
            },
            self.temp_lookup_value_K: {
                self.temp_lookup_value_C: (0, 1, -273.15),
                self.temp_lookup_value_F: (-273.15, 9/5, 32),
                self.temp_lookup_value_K: (0, 1, 0),
                self.temp_lookup_value_R: (0, 9/5, 0)
            },
            self.temp_lookup_value_R: {
                self.temp_lookup_value_C: (-491.67, 5/9, 0),
                self.temp_lookup_value_F: (0, 1, -459.67),
                self.temp_lookup_value_K: (0, 5/9, 0),
                self.temp_lookup_value_R: (0, 1, 0)
            }
        }
        self.dict_temp = {
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

        self.dict_length_convert = {
            self.length_lookup_value_mm: {
                self.length_lookup_value_mm: 1,
                self.length_lookup_value_cm: 1 / 10,
                self.length_lookup_value_dm: 1 / 100,
                self.length_lookup_value_m: 1 / 1000,
                self.length_lookup_value_mil: 39.3701,
                self.length_lookup_value_inch: 1 / 25.4,
                self.length_lookup_value_foot: 1 / 304.8,
                self.length_lookup_value_mile: 1 / (304.8 * 5280)
            },
            self.length_lookup_value_cm: {
                self.length_lookup_value_mm: 10,
                self.length_lookup_value_cm: 1,
                self.length_lookup_value_dm: 1 / 10,
                self.length_lookup_value_m: 1 / 100,
                self.length_lookup_value_mil: 393.701,
                self.length_lookup_value_inch: 1 / 2.54,
                self.length_lookup_value_foot: 1 / 30.48,
                self.length_lookup_value_mile: 1 / (30.48 * 5280)
            },
            self.length_lookup_value_dm: {
                self.length_lookup_value_mm: 100,
                self.length_lookup_value_cm: 10,
                self.length_lookup_value_dm: 1,
                self.length_lookup_value_m: 1 / 10,
                self.length_lookup_value_mil: 3937.01,
                self.length_lookup_value_inch: 1 / .254,
                self.length_lookup_value_foot: 1 / 3.048,
                self.length_lookup_value_mile: 1 / (3.048 * 5280)
            },
            self.length_lookup_value_m: {
                self.length_lookup_value_mm: 1000,
                self.length_lookup_value_cm: 100,
                self.length_lookup_value_dm: 10,
                self.length_lookup_value_m: 1,
                self.length_lookup_value_mil: 39370.1,
                self.length_lookup_value_inch: 1 / 0.0254,
                self.length_lookup_value_foot: 1 / 0.3048,
                self.length_lookup_value_mile: 1 / (.3048 * 5280)
            },
            self.length_lookup_value_mil: {
                self.length_lookup_value_mm: 0.0254,
                self.length_lookup_value_cm: 0.00254,
                self.length_lookup_value_dm: 0.000254,
                self.length_lookup_value_m: 0.0000254,
                self.length_lookup_value_mil: 1,
                self.length_lookup_value_inch: 0.001,
                self.length_lookup_value_foot: 1 / 12000,
                self.length_lookup_value_mile: 1 / (12000 * 5280)
            },
            self.length_lookup_value_inch: {
                self.length_lookup_value_mm: 25.4,
                self.length_lookup_value_cm: 2.54,
                self.length_lookup_value_dm: 0.254,
                self.length_lookup_value_m: 0.0254,
                self.length_lookup_value_mil: 1000,
                self.length_lookup_value_inch: 1,
                self.length_lookup_value_foot: 1 / 12,
                self.length_lookup_value_mile: 1 / (12 * 5280)
            },
            self.length_lookup_value_foot: {
                self.length_lookup_value_mm: 304.8,
                self.length_lookup_value_cm: 30.48,
                self.length_lookup_value_dm: 3.048,
                self.length_lookup_value_m: 0.3048,
                self.length_lookup_value_mil: 12000,
                self.length_lookup_value_inch: 12,
                self.length_lookup_value_foot: 1,
                self.length_lookup_value_mile: 1 / 5280
            },
            self.length_lookup_value_mile: {
                self.length_lookup_value_mm: (304.8 * 5280),
                self.length_lookup_value_cm: (30.48 * 5280),
                self.length_lookup_value_dm: (3.048 * 5280),
                self.length_lookup_value_m: (0.3048 * 5280),
                self.length_lookup_value_mil: (12000 * 5280),
                self.length_lookup_value_inch: (12 * 5280),
                self.length_lookup_value_foot: 5280,
                self.length_lookup_value_mile: 1
            }
        }
        self.dict_length = {
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
            'feet': self.length_lookup_value_foot,

            'mi': self.length_lookup_value_mile,
            'mile': self.length_lookup_value_mile,
            'miles': self.length_lookup_value_mile
        }

        self.speed_lookup_value_mps = 0  # meters per second
        self.speed_lookup_value_kmh = 1  # kilometers per hour
        self.speed_lookup_value_fps = 2  # feet per second
        self.speed_lookup_value_fph = 3  # feet per second
        self.speed_lookup_value_mph = 4  # miles per hour
        self.speed_lookup_value_knots = 5  # knots

        self.dict_speed_convert = {
            self.speed_lookup_value_mps: {
                self.speed_lookup_value_mps: 1,
                self.speed_lookup_value_kmh: 3.6,
                self.speed_lookup_value_fps: 100 / 30.48,
                self.speed_lookup_value_fph: 100 / 30.48 * 3600,
                self.speed_lookup_value_mph: 100 / 160934.4 * 3600,
                self.speed_lookup_value_knots: 1.943844
            },
            self. speed_lookup_value_kmh: {
                self.speed_lookup_value_mps: 1 / 3.6,
                self.speed_lookup_value_kmh: 1,
                self.speed_lookup_value_fps: 100000 / 109725,
                self.speed_lookup_value_fph: 100000 / 30.48,
                self.speed_lookup_value_mph: 100000 / 160934.4,
                self.speed_lookup_value_knots: 1 / 3.6 * 1.943844
            },
            self.speed_lookup_value_fps: {
                self.speed_lookup_value_mps: 0.3048,
                self.speed_lookup_value_kmh: 1.09728,
                self.speed_lookup_value_fps: 1,
                self.speed_lookup_value_fph: 3600,
                self.speed_lookup_value_mph: 3600 / 5280,
                self.speed_lookup_value_knots: 0.3048 * 1.943844
            },
            self.speed_lookup_value_fph: {
                self.speed_lookup_value_mps: 30.48 / (3600 * 100),
                self.speed_lookup_value_kmh: 30.48 / (100 * 1000),
                self.speed_lookup_value_fps: 1 / 3600,
                self.speed_lookup_value_fph: 1,
                self.speed_lookup_value_mph: 1 / 5280,
                self.speed_lookup_value_knots: 30.48 / (3600 * 100) * 1.943844
            },
            self.speed_lookup_value_mph: {
                self.speed_lookup_value_mps: 0.44704,
                self.speed_lookup_value_kmh: 1.609344,
                self.speed_lookup_value_fps: 5280 / 3600,
                self.speed_lookup_value_fph: 5280,
                self.speed_lookup_value_mph: 1,
                self.speed_lookup_value_knots: 0.44704 * 1.943844
            },
            self.speed_lookup_value_knots: {
                self.speed_lookup_value_mps: 1 / 1.943844,
                self.speed_lookup_value_kmh: 1 / 1.943844 * 3.6,
                self.speed_lookup_value_fps: 1 / 1.943844 * 1 / 0.3048,
                self.speed_lookup_value_fph: 1 / 1.943844 * 3600 / 0.3048,
                self.speed_lookup_value_mph: 1 / 1.943844 * 3.6 / 1.609344,
                self.speed_lookup_value_knots: 1
            }
        }
        self.dict_speed = {
            'mps': self.speed_lookup_value_mps,
            'meters/s': self.speed_lookup_value_mps,
            'meters/sec': self.speed_lookup_value_mps,
            'meters/second': self.speed_lookup_value_mps,
            'm/s': self.speed_lookup_value_mps,

            'kmh': self.speed_lookup_value_kmh,
            'kilometers/s': self.speed_lookup_value_kmh,
            'kilometers/sec': self.speed_lookup_value_mps,
            'kilometers/second': self.speed_lookup_value_mps,
            'km/s': self.speed_lookup_value_kmh,

            'fps': self.speed_lookup_value_fps,
            'feet/s': self.speed_lookup_value_fps,
            'feet/sec': self.speed_lookup_value_fps,
            'feet/second': self.speed_lookup_value_fps,
            'ft/s': self.speed_lookup_value_fps,

            'fph': self.speed_lookup_value_fph,
            'feet/h': self.speed_lookup_value_fph,
            'feet/hr': self.speed_lookup_value_fph,
            'feet/hour': self.speed_lookup_value_fph,
            'foot/h': self.speed_lookup_value_fph,
            'foot/hr': self.speed_lookup_value_fph,
            'foot/hour': self.speed_lookup_value_fph,
            'ft/hr': self.speed_lookup_value_fph,
            'ft/h': self.speed_lookup_value_fph,

            'mph': self.speed_lookup_value_mph,
            'm.p.h.': self.speed_lookup_value_mph,
            'MPH': self.speed_lookup_value_mph,
            'mi/hour': self.speed_lookup_value_mph,

            'kn': self.speed_lookup_value_knots,
            'kt': self.speed_lookup_value_knots,
            'knot': self.speed_lookup_value_knots,
            'knots': self.speed_lookup_value_knots,
        }

    def temp_convert(self, value, input_units, output_units):
        """
        Converts temperature from input units to output units
        :param value: value to be converted
        :param input_units: input units C/F/K/R
        :param output_units: output units C/F/K/R
        :return: converted temperature
        """
        try:
            conversion = self.dict_temp_convert[self.dict_temp[input_units]][self.dict_temp[output_units]]
            output = (value + conversion[0]) * conversion[1] + conversion[2]

        except KeyError:
            return "error"
        return output

    def speed_convert(self, value, input_units, output_units):
        """
        Converts speed from input units to output units
        :param value: value to be converted
        :param input_units: input units ex. m/s, fts, mph, etc...
        :param output_units: output units ex. m/s, fts, mph, etc...
        :return: converted temperature
        """
        try:
            conversion = self.dict_speed_convert[self.dict_speed[input_units]][self.dict_speed[output_units]]
            output = (value * conversion)
        except KeyError:
            return "error"
        return output

    def length_convert(self, value, input_units, output_units):
        """
        Converts length from input units to output units
        :param value: value to be converted
        :param input_units: input units
        :param output_units: output units
        :return: converted distance
        """
        try:
            conversion = self.dict_length_convert[self.dict_length[input_units]][self.dict_length[output_units]]
            output = (value * conversion)
        except KeyError:
            return "error"
        return output

    @staticmethod
    def temp_test(val):
        print('Test calc, C/F/K/R')
        list_ = {'C', 'F', 'K', 'R'}
        for x in list_:
            for y in list_:
                _val = app.temp_convert(val, x, y)
                print(f'{val} {x} converted to {y}: {_val}')

    @staticmethod
    def length_test(val):
        list_ = {'mm', 'cm', 'dm', 'm', 'mil', 'inch', 'ft', 'mile'}
        for x in list_:
            for y in list_:
                _val = app.length_convert(val, x, y)
                print(f'{val} {x} converted to {y}: {_val}')\

    @staticmethod
    def speed_test(val):
        list_ = {'mps', 'kmh', 'fps', 'fph', 'mph', 'knots'}
        for x in list_:
            for y in list_:
                _val = app.speed_convert(val, x, y)
                print(f'{val} {x} converted to {y}: {_val}')


if __name__ == "__main__":
    app = UnitConvert()
    app.speed_test(10)
    # app.length_test(10)
