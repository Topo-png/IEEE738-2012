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
import main as ieee738

app = ieee738.IEEE738()
#
demo = True  # disable to allow selection of config/conductors

path_config = 'Sample/config-sample.xlsx'  # location of configuration file
path_conductor = 'Sample/Conductor_Prop-Sample.xlsx'  # location of conductor file


def runTest():
    config_list = app.import_config(path_config, sheet_name='config')
    conductor_list, conductor_spec_temp_list = \
        app.import_conductor(path_conductor, ['conductors', 'conductor spec'])

    df_config = select_config(config_list)
    df_conductor, df_spec = select_conductor(conductor_list, conductor_spec_temp_list)
    app.true_to_standard = False  # mimics original PJM spreadsheet performance

    df_normal, df_emergency, df_load = app.c_reporting(df_conductor, df_spec, df_config)

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

    app.export_excel(df_normal, df_emergency, df_load, df_config, 'export_test')

    print('Configuration info:', df_config)
    print('Conductor info:', df_conductor)
    print('Daytime normal rating', df_normal_day)
    print('Nighttime normal rating', df_normal_night)
    print('Daytime emergency rating', df_emergency_day)
    print('Nighttime emergency rating', df_emergency_night)
    print('Daytime load dump rating', df_load_day)
    print('Nighttime load dump rating',  df_load_night)

    return None


def select_config(df_config):
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
        app.select_config(df_config)
    except IndexError:
        print('Please select a valid configuration')
        app.select_config(df_config)
    except TypeError:
        print('Please select a valid configuration')
        app.select_config(df_config)
    except ValueError:
        print('Please select a valid configuration')
        app.select_config(df_config)
    return config


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


if __name__ == "__main__":
    runTest()
