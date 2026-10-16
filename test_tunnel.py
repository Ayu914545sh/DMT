import requests
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings("ignore")

s = requests.Session()
URL = 'https://tanks-clearly-paid-pictures.trycloudflare.com'
print("Tunnel URL:", URL)

# 1. Login
r = s.get(f'{URL}/login/')
soup = BeautifulSoup(r.text, 'html.parser')
csrf = soup.find('input', {'name': 'csrfmiddlewaretoken'})['value']

r2 = s.post(f'{URL}/login/', data={
    'username': 'datamanager',
    'password': 'manager123',
    'csrfmiddlewaretoken': csrf
}, headers={'Referer': f'{URL}/login/'})

print('Login URL after post:', r2.url)
if r2.url != f'{URL}/':
    print("WARNING: Login may have failed or redirected.")

# 2. Upload
r5 = s.get(f'{URL}/import/')
soup = BeautifulSoup(r5.text, 'html.parser')
csrf = soup.find('input', {'name': 'csrfmiddlewaretoken'})['value']
files = {'data_file': ('sample_mca.csv', open('sample_data/sample_mca.csv', 'rb'), 'text/csv')}
r6 = s.post(f'{URL}/import/', data={
    'dataset_type': 'MCA',
    'header_row': '1',
    'csv_delimiter': ',',
    'csv_encoding': 'utf-8',
    'csrfmiddlewaretoken': csrf
}, files=files, headers={'Referer': f'{URL}/import/'})
print('Upload redirected to:', r6.url)
if '/map/' in r6.url:
    print('Upload Success: Mapped!')
else:
    print('Upload Error/CSRF Failed:', r6.status_code)

# 3. Export
r7 = s.get(f'{URL}/exports/')
soup = BeautifulSoup(r7.text, 'html.parser')
csrf = soup.find('input', {'name': 'csrfmiddlewaretoken'})['value']
r8 = s.post(f'{URL}/exports/', data={
    'dataset_type': 'MCA',
    'file_format': 'CSV',
    'state_filter': '',
    'status_filter': '',
    'csrfmiddlewaretoken': csrf
}, headers={'Referer': f'{URL}/exports/'})
print('Export submission status:', r8.status_code)

# 4. Logout
r9 = s.post(f'{URL}/logout/', data={'csrfmiddlewaretoken': csrf}, headers={'Referer': f'{URL}/'})
print('Logout status:', r9.status_code)
