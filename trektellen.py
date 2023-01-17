import xlrd3 as xlrd
import xlsxwriter
import math
import datetime
import sys


class BenFiles:

    def __init__(self, input_filename):
        self.filename = input_filename

    @property
    def headers(self):
        return ['Common Name',   'Genus', 'Species',
                '# S', '# N', 'RES', 'Spec. Comments', 'Location Name',
                'Latitude', 'Longitude', 'DATE', 'TIME',
                'State/Province', 'Country Code', 'Protocol', '#Obs', 'OBS.', 'DURATION',
                'All observations reported?',
                'Effort Distance Miles', 'CLOUDS', 'VISIB', 'WIND', 'Effort area acres',
                'Submission Comments']

    def date(self, cell_value):
        try:
            return datetime.datetime(*xlrd.xldate_as_tuple(cell_value, datemode=0))
        except TypeError:
            for key in ['/', '-']:
                if key in cell_value:
                    splitted = cell_value.split(key)
                    return datetime.datetime(year=int(splitted[2]),
                                             month=int(splitted[1]),
                                             day=int(splitted[0]))
            assert type(cell_value) == datetime
            return 'WTF'

    def get_data(self):
        book = xlrd.open_workbook(self.filename)
        sheet = book.sheet_by_index(0)
        informations = []
        for nb_row in range(sheet.nrows):
            if nb_row == 0:
                continue
            obs = dict()
            for index, header in enumerate(self.headers):
                cell_value = sheet.cell_value(rowx=nb_row, colx=index)
                if header == 'DATE' and cell_value != '':
                    cell_value = self.date(cell_value)
                obs[header] = cell_value
            informations.append(obs)

        return informations

    def show_data(self):
        for obs in self.get_data():
            print(obs)


class TrekTellenFile:

    def __init__(self, out_filename):
        self.all_dates = {}
        self.workbook = xlsxwriter.Workbook(out_filename)
        self.header_sheet = self.workbook.add_worksheet("Header")
        self.species_sheet = self.workbook.add_worksheet("Species")
        self.current_header_line = 1
        self.current_species_line = 1
        for index, header in enumerate(self.general_headers):
            self.header_sheet.write(0, index, header)
        for index, header in enumerate(self.species_headers):
            self.species_sheet.write(0, index, header)

    def close(self):
        self.workbook.close()

    @property
    def general_headers(self):
        return [
            'id', 'siteid', 'date', 'start', 'end', 'observers', 'weather', 'windspeed_bfr',
            'wind_ms', 'winddirection', 'cloudcover', 'cloudheight', 'precipitation',
            'perc_duration', 'visibility', 'temperature', 'observersactive', 'observerspresent',
            'counttype', 'remarks'
        ]

    @property
    def species_headers(self):
        return ['date', 'timestamp', 'countid', 'siteid', 'speciesid', 'speciesname',
                'direction1', 'direction2', 'local', 'remarkable', 'remarkablelocal',
                'age', 'sex', 'plumage', 'remark', 'location', 'migtype', 'counttype',
                'year', 'yday', 'exactdirection1',
                'exactdirection2', 'groupid', 'submitted']

    def start(self, start):
        date = str(start).split('.')
        hours, minutes = int(date[0]), int(date[1])
        return f'{hours:02d}:{minutes:02d}'

    def end(self, start, duration):
        date = str(start).split('.')
        hours, minutes = int(date[0]), int(date[1])
        total_minutes = minutes + duration
        hours = math.floor((hours * 60 + total_minutes) / 60)
        minutes = int(total_minutes % 60)
        return f'{hours:02d}:{minutes:02d}'

    def weather(self, obs):
        """
        You can add the full coulds text to the weather text field.
        """
        return f'Clouds : {obs["CLOUDS"]}  / Visibility : {obs["VISIB"]} / Wind : {obs["WIND"]}'

    def split_data(self, ben_info):

        if '>' not in ben_info:
            return int(ben_info)
        splitted = ben_info.split('>')
        if len(splitted) == 1:
            return int(splitted[0])
        return (int(splitted[0]) + int(splitted[1])) / 2

    def cloud_cover(self, cloud):
        """
        Couldcover can be a percentage, but is should be one single value.
        So in the vast of 20>100 is is probably best to use 55.
        """
        if type(cloud) == float:
            return int(cloud)
        cloud = cloud.split(';')[0]

        return self.split_data(cloud)

    def precipitation(self, obs):
        """
        In Precipitation there should be a value like, rain, hail.
        """
        clouds = obs['CLOUDS']
        if type(clouds) == float:
            return ""
        if ";" in clouds:
            return clouds.split(';')[-1]
        return ""

    def wind_speed(self, wind):
        """
        Winder speed is an integer value between 0 and 12
        """
        if wind == '':
            return 0
        wind_speed = wind.rstrip().split(' ')[-1]
        wind_speed = wind_speed.split('A')[-1]
        return int(wind_speed)

    def wind_direction(self, wind):
        """
            Wind direction : like WSW
        """
        wind_dir = wind.split(' ')[0]
        return wind_dir

    def visibility(self, visibility):
        """
        Visibility: single integer field,  infinity just choose something: 25000 is okay.
        Again if you have an value like 8 > 4 km choose 6 as the value and add this text to weather field.
        """
        if visibility == '∞':
            return 25000

        multiply = False
        if 'KM' in visibility:
            multiply = True
            visibility = visibility.split('KM')[0]

        new_visib = self.split_data(visibility.replace('∞', '25000'))
        if multiply:
            return new_visib * 1000
        return new_visib

    def populate_header(self, obs):
        infos = {}
        for header in self.general_headers:
            infos[header] = ''
        infos['id'] = 'COUNT_ID'
        infos['siteid'] = 'CHOCO_ID'
        infos['date'] = obs['DATE'].strftime('%d/%m/%Y')
        infos['start'] = self.start(obs['TIME'])
        infos['end'] = self.end(obs['TIME'], obs['DURATION'])
        infos['observers'] = obs['OBS.']
        infos['windspeed_bfr'] = self.wind_speed(obs['WIND'])
        infos['wind_ms'] = self.wind_speed(obs['WIND'])
        infos['winddirection'] = self.wind_direction(obs['WIND'])
        infos['weather'] = self.weather(obs)
        infos['precipitation'] = self.precipitation(obs)
        infos['cloudcover'] = self.cloud_cover(obs['CLOUDS'])
        infos['visibility'] = self.visibility(obs['VISIB'])
        infos['observersactive'] = len(obs['OBS.'].split(' '))
        infos['observerspresent'] = len(obs['OBS.'].split(' '))
        infos['counttype'] = 'seawatch'

        return infos

    def add_header(self, obs):
        if obs['DATE'] in self.all_dates:
            return
        infos = self.populate_header(obs)
        for index, header in enumerate(self.general_headers):
            self.header_sheet.write(self.current_header_line, index, infos[header])
        self.all_dates[obs['DATE']] = True
        self.current_header_line += 1

    def populate_species(self, obs):
        infos = {}
        for header in self.species_headers:
            infos[header] = ''

        infos['date'] = obs['DATE'].strftime('%d/%m/%Y')
        infos['timestamp'] = self.start(obs['TIME'])
        infos['countid'] = 'COUNT_ID'
        infos['siteid'] = 'CHOCO_ID'
        infos['speciesid'] = obs['Common Name']
        infos['speciesname'] = obs['Common Name']
        infos['direction1'] = obs['# S']
        infos['direction2'] = obs['# N']
        infos['local'] = obs['RES']
        infos['remark'] = obs['Spec. Comments']
        infos['year'] = obs['DATE'].strftime('%Y')
        infos['yday'] = obs['DATE'].timetuple().tm_yday
        infos['groupid'] = 'GROUP_ID'
        infos['submitted'] = 'SUBMITTED_TIME'
        return infos

    def add_data(self, obs):
        if obs['OBS.'] == '':
            return
        self.add_header(obs)
        infos = self.populate_species(obs)
        for index, header in enumerate(self.species_headers):
            self.species_sheet.write(self.current_species_line, index, infos[header])
        self.current_species_line += 1


n = len(sys.argv)
assert n == 2
filename = sys.argv[1]

ben_data = BenFiles(filename).get_data()

print('TREKTELLEN HEADER FILE')
trek = TrekTellenFile('trektellen_out.xls')
for line in ben_data:
    print('LINE', line)
    trek.add_data(line)

trek.close()
