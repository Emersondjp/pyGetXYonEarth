
from geopy.geocoders import GoogleV3
geolocator = GoogleV3()
location = geolocator.geocode("175 5th Avenue NYC")
print(location.address)
print((location.latitude, location.longitude))
print(location.raw)

