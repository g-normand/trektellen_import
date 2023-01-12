import xlrd3 as xlrd
import math
import datetime


class BenFiles:

    def __init__(self, filename):
        self.filename = filename

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
            print('CELL', cell_value)
            return cell_value

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


class TrekTellenHeaderFile:

    @property
    def headers(self):
        return [
            'id', 'siteid', 'date', 'start', 'end', 'observers', 'weather', 'windspeed_bfr', 'wind_ms',
            'winddirection', 'cloudcover', 'cloudheight', 'precipitation', 'perc_duration', 'visibility',
            'temperature', 'observersactive', 'observerspresent', 'counttype', 'remarks'
        ]

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

    def populate(self, obs):
        infos = {}
        for header in self.headers:
            infos[header] = ''
        infos['siteid'] = 'CHOCO_ID'
        infos['date'] = obs['DATE']
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

    def show_data(self, obs):
        if obs['OBS.'] == '':
            return
        infos = self.populate(obs)
        for header in self.headers:
            print(header, infos[header])


print('BEN FILE')
# ben_data = BenFiles("test.xls").get_data()
ben_data = BenFiles("2018.xlsx").get_data()

print('TREKTELLEN HEADER FILE')
for line in ben_data:
    print('LINE', line)
    TrekTellenHeaderFile().show_data(line)

