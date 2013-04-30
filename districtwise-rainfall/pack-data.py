

from string import strip, split
import sys, os
from xlwt import Workbook


# the line numbers on which the 5 yr data are located,
y_linenos = [18, 20, 22, 24, 26]
# the data which is missing is denoted by a blank in the data file.
# Only way to determine each months data is then to look at the specific
# positions (since all files have same format). The month's avg,diff rainfall
# position in a line are given by this list.
m_pos = [(11+13*i, 17+13*i) for i in range(12)]


def get_data(fname):
    y_dat = {}
    with open(fname) as f:
        lines = f.readlines()
        for i in y_linenos:
            #d = map(float, split(l[i]))[1:] doesn't work because of blank/missing data
            l = lines[i]
            y = strip(l[:5])
            m_dat = []
            for j in range(12):
                # try-except clauses are to handle blanks.
                m_avg, m_diff = 0, 0
                try: m_avg = float(l[m_pos[j][0]-6:m_pos[j][0]]) 
                except: m_avg = -1
                try: m_diff = float(l[m_pos[j][1]-6:m_pos[j][1]])
                except: m_diff = 0
                m_dat.append([m_avg, m_diff])
            y_dat[y] = m_dat       
    return y_dat        



months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
        'Sep', 'Oct', 'Nov', 'Dec']

def write_sheet(wb, sheetname, dat):
    ws = wb.add_sheet(sheetname)
    ws.row(0).write(0, sheetname)
    for i in range(12):
        ws.row(0).write(i+1, months[i])
    
    y = sorted(dat.keys())
    for i in range(len(dat)):
        m_dat = dat[y[i]]
        ws.row(2*i+1).write(0, y[i]+'-avg')
        ws.row(2*i+2).write(0, y[i]+'-diff')
        for j in range(12):
            #print m_dat[j]
            ws.row(2*i+1).write(j+1, m_dat[j][0])
            ws.row(2*i+2).write(j+1, m_dat[j][1])
            




DEBUG = False

def pack_state(d):
    if DEBUG: print 'packing for', d, '> '
    for rs, ds, fs in os.walk(d):
        fs = [f for f in fs if f.endswith('.txt')]
        if DEBUG: print fs
        if fs==[]: return False
        book = Workbook()
        for f in fs:
            if f.endswith('.txt'):
                sheetname = f[:-4]
                if DEBUG: print sheetname,
                r_dat = get_data(d+'/'+f)
                write_sheet(book, sheetname, r_dat)
        book.save(d+'.xls')
        if DEBUG: print 
        return True

def pack_all():
    import zipfile, zlib
    arch_name = 'india-districtwise-rain-data-2007-11'
    with zipfile.ZipFile(arch_name+'.zip', 'w', 
            zipfile.ZIP_DEFLATED) as zippack:
        for root, dirs, files in os.walk('.'):
            for d in dirs:
                if pack_state(d): zippack.write(d+'.xls')


#process_state('orissa')
pack_all()


