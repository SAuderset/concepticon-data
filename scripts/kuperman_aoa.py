import xlrd
from cldfcatalog import Config 
from csvw.dsv import UnicodeDictReader 
from collections import defaultdict

# get mappings etc.
repos = Config.from_file().get_clone('concepticon') 
paths = {p.stem.split('-')[1]: p for p in repos.joinpath(
    'mappings').glob('map-*.tsv')}
mappings = {} 
for language, path in paths.items(): 
    mappings[language] = defaultdict(set) 
    with UnicodeDictReader(path, delimiter='\t') as reader: 
        for line in reader: 
            gloss = line['GLOSS'].split('///')[1] 
            mappings[language][gloss].add( 
                (line['ID'], int(line['PRIORITY'])))

# get xlsx file
xlfile = xlrd.open_workbook('raw/13428_2013_348_MOESM1_ESM.xlsx')
# retrieve sheet
sheet = xlfile.sheet_by_index(0)

# prepare mapping
data = defaultdict(list)

# get data and map them if possible
for i in range(1, sheet.nrows):
    gloss, occtot, occnum, freq, ratm, rats, dun = sheet.row_values(i)
    if gloss in mappings['en']: 
        best_match, priority = sorted( 
            mappings['en'][gloss], 
            key=lambda x: x[1])[0] 
        data[best_match] += [[ 
            str(i), 
            gloss, 
            str(occtot),
            str(occnum),
            '{0:.2f}'.format(freq),
            '{0:.2f}'.format(ratm),
            '{0:.2f}'.format(rats),
            '{0:.2f}'.format(dun),
            best_match, 
            priority]]

with open('kuperman_aoa.tsv', 'w') as f: 
    f.write('\t'.join(["CONCEPTICON_ID",
            "AOA_WORD",
            "OCCUR_TOTAL",
            "OCCUR_NUM",
            "FREQ_PM",
            "RATING_MEAN",
            "RATING_SD",
            "DUNNO",
            "LINE_IN_SOURCE"])+'\n')
    for key, lines in data.items(): 
        best_line = sorted(lines, key=lambda x: x[-1])[0] 
        best_line[-1] = str(best_line[-1]) 
        f.write('\t'.join([
            best_line[-2], # concepticon id
            best_line[1],
            best_line[2], 
            best_line[3],
            best_line[4],
            best_line[5],
            best_line[6],
            best_line[7],
            best_line[0]
            ])+'\n')


print('Found {0} direct matches in data.'.format(len(data)))
