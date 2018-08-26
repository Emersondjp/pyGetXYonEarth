
#coding=utf-8
import xlrd, xlwt, re, os, sys, time
from xlutils.copy import copy
import geopy
import googlemaps

pxy = {'http':'111.13.109.27:80'}
FileName = u'2010国内企业-arcgis.xls'

#geolocator = geopy.Nominatim(timeout=120)
#geolocator = geopy.GoogleV3(proxies=pxy)
#geolocator = googlemaps.Client(api_key='HkoHlk2LqDgNgMf5ufuByLnIRneVy6oj')
geolocator = geopy.Baidu(api_key='HkoHlk2LqDgNgMf5ufuByLnIRneVy6oj')

def getXY( cityName ):
    print( cityName )
    str_len = len( cityName )
    location = None
    i = -1
    while location == None:
        if i+str_len == 0 : return (-1, -1, "NULL")
        print( cityName[0:i] )
        try :
            location = geolocator.geocode( cityName[0:i], timeout=None )
            i = i - 1
        except geopy.exc.GeocoderTimedOut as e:
            print("Timed out. Wait for 10 seconds.")
            time.sleep(10)
    return (location.latitude, location.longitude, cityName[0:i])

rb = xlrd.open_workbook( FileName )
r_sheet = rb.sheet_by_index(1)

nrows = r_sheet.nrows
wb = xlwt.Workbook()

w_sheet = wb.add_sheet(u'地址经纬度坐标')

w_sheet.write(0,0,u'通讯地址')
w_sheet.write(0,1,u'经度')
w_sheet.write(0,2,u'维度')
w_sheet.write(0,3,u'经纬度')
w_sheet.write(0,4,u'匹配地址')

for rowNo in range(1, nrows):
    if (r_sheet.cell( rowNo, 8 ).value).strip() == "" : continue
    if unicode(r_sheet.cell( rowNo, 8 ).value).strip() == u"-" : continue

    addr = r_sheet.cell( rowNo, 8 ).value
    print(unicode(addr))
    x, y, loc = getXY( unicode(addr).strip() )
    w_sheet.write( rowNo, 0, addr )
    w_sheet.write( rowNo, 1, x )
    w_sheet.write( rowNo, 2, y )
    w_sheet.write( rowNo, 3, "(%s, %s)" % (x, y) )
    w_sheet.write( rowNo, 4, "%s" % loc )
#    print( "Addr: %s, X: %s, Y: %s\n" % ( addr, x, y ) )
    wb.save( 'out.xls' )
time.sleep(5)

wb.save( 'out.xls' )

