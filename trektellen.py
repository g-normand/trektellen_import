import xlrd3 as xlrd
import xlsxwriter
import math
import datetime
import sys

CHOCOLATERA_ID = 2788
SPECIES_ID = dict()
SPECIES_ID['ACTITIS MACULARIA'] = 186
SPECIES_ID['ANAS BAHAMENSIS'] = 554
SPECIES_ID['ANAS DISCORS'] = 65
SPECIES_ID['ANOUS STOLIDUS'] = 708
SPECIES_ID['ARDEA ALBA'] = 28
SPECIES_ID['ARDEA COCOI'] = 3146
SPECIES_ID['ARDENNA CREATOPUS'] = 494
SPECIES_ID['ARDENNA GRISEA'] = 13
SPECIES_ID['ARENARIA INTERPRES'] = 187
SPECIES_ID['BUBULCUS IBIS'] = 26
SPECIES_ID['CALIDRIS ALBA'] = 152
SPECIES_ID['CALIDRIS MAURI'] = 650
SPECIES_ID['CALIDRIS MINUTILLA'] = 652
SPECIES_ID['CALIDRIS VIRGATA'] = 2010
SPECIES_ID['CHARADRIUS SEMIPALMATUS'] = 636
SPECIES_ID['CHARADRIUS VOCIFERUS'] = 637
SPECIES_ID['CHLIDONIAS NIGER'] = 222
SPECIES_ID['CHROICOCEPH. CIRROCEPH.'] = 1024
SPECIES_ID['CREAGRUS FURCATUS'] = 2013
SPECIES_ID['DENDROCYGNA AUTUMNALIS'] = 1297
SPECIES_ID['EGRETA ALBA'] = 28
SPECIES_ID['EGRETTA ALBA'] = 28
SPECIES_ID['EGRETTA THULA'] = 521
SPECIES_ID['EUDOCIMUS ALBUS'] = 531
SPECIES_ID['FALCO PEREGRINUS'] = 115
SPECIES_ID['FREGATA MAGNIFICENS'] = 512
SPECIES_ID['GELOCHELIDON NILOTICA'] = 213
SPECIES_ID['HAEMATOPUS PALLIATUS'] = 626
SPECIES_ID['HYDROBATES HORBYI'] = 1957
SPECIES_ID['LAROSTERNA INCA'] = 3221
SPECIES_ID['LARUS DOMINICANUS'] = 1616
SPECIES_ID['LEUCOPHAEUS ATRICILLA'] = 675
SPECIES_ID['LEUCOPHAEUS ATRICLLA'] = 675
SPECIES_ID['LEUCOPHAEUS MODESTUS'] = 3218
SPECIES_ID['LEUCOPHAEUS PIPIXCAN'] = 464
SPECIES_ID['NUMENIUS PHAEOPUS'] = 173
SPECIES_ID['NYCTANASSA VIOLACEA'] = 1350
SPECIES_ID['OCEANITES  GRACILIS'] = 3124
SPECIES_ID['OCEANODROMA sp'] = 500
SPECIES_ID['OCEANODROMA TETHYS'] = 1952
SPECIES_ID['PELECANUS OCCIDENTALIS'] = 509
SPECIES_ID['PELECANUS OCCIDENTALIS M'] = 509
SPECIES_ID['PELECANUS THAGUS'] = 4683
SPECIES_ID['PHAETHON AETHEREUS'] = 18
SPECIES_ID['PHALACROC. BOUGAINVILLII'] = 3149
SPECIES_ID['PHALACROC. BRASILIANUS'] = 1365
SPECIES_ID['PHALAROPUS FULICARIUS'] = 190
SPECIES_ID['PHALAROPUS LOBATUS'] = 189
SPECIES_ID['PHALAROPUS SPEC'] = 1092
SPECIES_ID['PHALAROPUS TRICOLOR'] = 188
SPECIES_ID['PHOEBASTRIA IRRORATA'] = 3126
SPECIES_ID['PLUVIALIS SQUATAROLA'] = 147
SPECIES_ID['PTERODROMA PHAEOPYGIA'] = 3128
SPECIES_ID['PUFFINUS SPEC'] = 229
SPECIES_ID['PUFFINUS SUBALARIS'] = 3131
SPECIES_ID['RYNCHOPS NIGER'] = 710
SPECIES_ID['SRENA SPEC'] = 705
SPECIES_ID['STERCORARIUS LONGICAUDUS'] = 193
SPECIES_ID['STERCORARIUS PARASITICUS'] = 192
SPECIES_ID['STERCORARIUS POMARINUS'] = 191
SPECIES_ID['STERCORARIUS SPEC'] = 455
SPECIES_ID['STERNA HIRUNDINACEA'] = 3220
SPECIES_ID['STERNA HIRUNDO'] = 217
SPECIES_ID['STERNA PARADISAEA'] = 218
SPECIES_ID['STERNA SPEC'] = 705
SPECIES_ID['SULA GRANTI'] = 3151
SPECIES_ID['SULA LEUCOGASTER'] = 501
SPECIES_ID['SULA NEBOUXII'] = 1979
SPECIES_ID['SULA SULA'] = 1980
SPECIES_ID['SULA VARIEGATA'] = 3150
SPECIES_ID['SYULA VARIEGATA'] = 3150
SPECIES_ID['THALASSEUS ELEGANS'] = 696
SPECIES_ID['THALASSEUS ELEGANT'] = 696
SPECIES_ID['THALASSEUS MAXIMUS'] = 692
SPECIES_ID['THALASSEUS SANDVCENSIS'] = 215
SPECIES_ID['THALASSEUS SANDVICENSIS'] = 215
SPECIES_ID['THALASSEUS SPEC'] = 4686
SPECIES_ID['TRINGA INCANA'] = 670
SPECIES_ID['TRINGA SEMIPALMATA'] = 671
SPECIES_ID['WADER SPEC'] = 672
SPECIES_ID['XEMA SABINI'] = 198


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
            # US format need to be changed
            date_tuple = xlrd.xldate_as_tuple(cell_value, datemode=0)
            return datetime.datetime(year=date_tuple[0],
                                     month=date_tuple[2],
                                     day=date_tuple[1])
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

        Cloudcover should be a value between 0 and 8
        25% will be 2/8
        """
        if type(cloud) == float:
            return int(cloud / 12.5)
        cloud = cloud.split(';')[0]

        return int(self.split_data(cloud) / 12.5)

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
        infos['id'] = self.current_header_line
        infos['siteid'] = CHOCOLATERA_ID
        infos['date'] = obs['DATE'].strftime('%Y/%m/%d')
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
        infos['observersactive'] = len(obs['OBS.'].split(','))
        infos['observerspresent'] = len(obs['OBS.'].split(','))
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

    def remark_species(self, obs):
        remark = ''
        if obs['Genus'] != '':
            remark += f"Sub-species : {obs['Genus']}"

        if obs['Common Name'] == 'PELECANUS OCCIDENTALIS M':
            remark += "Sub-species : MURPHYI"

        remark += obs['Spec. Comments']

        return remark

    def populate_species(self, obs):
        infos = {}
        for header in self.species_headers:
            infos[header] = ''

        infos['date'] = obs['DATE'].strftime('%Y/%m/%d')
        infos['timestamp'] = self.start(obs['TIME'])
        infos['countid'] = self.current_header_line - 1
        infos['siteid'] = CHOCOLATERA_ID
        infos['speciesid'] = SPECIES_ID[obs['Common Name'].rstrip()]
        infos['speciesname'] = obs['Common Name']
        infos['direction1'] = obs['# S']
        infos['direction2'] = obs['# N']
        infos['local'] = obs['RES']
        infos['remark'] = self.remark_species(obs)
        infos['year'] = obs['DATE'].strftime('%Y')
        infos['yday'] = obs['DATE'].timetuple().tm_yday
        infos['groupid'] = ''
        infos['submitted'] = ''
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

trek = TrekTellenFile('trektellen_out.xls')
for line in ben_data:
    trek.add_data(line)

trek.close()
print('TREKTELLEN FILE OK')
