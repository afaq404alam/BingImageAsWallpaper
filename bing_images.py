import json
import os
import random
import re
import urllib.error
import urllib.request
import winreg

import requests

from tray_notify import balloon_tip

if os.name is 'nt':
    from ctypes import windll
else:
    raise Exception('Unsupported OS.')


class BingRandomImages:
    SPIF_UPDATEINIFILE = 0x01
    SPIF_SENDWININICHANGE = 0x02
    SPI_SETDESKWALLPAPER = 0x0014

    def __init__(self):
        self._random_image_file = 'random_bing_image.jpg'

    def set_wallpaper(self):
        servers = ['//az608707.vo.msecnd.net/files/', '//az619519.vo.msecnd.net/files/', '//az619822.vo.msecnd.net/files/']
        browse_data_url = 'https://www.bing.com/gallery/home/browsedata'
        image_details_url = 'https://www.bing.com/gallery/home/imagedetails/'

        bing_dir = r'C:\Users\{0}\Pictures\Bing'.format(os.getlogin())
        file_path = r'{0}\{1}'.format(bing_dir, "bing_random_image.jpg")

        if os.path.exists(file_path):
            os.remove(file_path)

        image_name, image_title = self.get_image_details(browse_data_url, image_details_url)

        self.download_image(servers, file_path, image_name)
        self.update_registry_values(file_path)
        self.refresh_desktop_wallpaper(file_path, image_desc=image_title)

    @staticmethod
    def update_registry_values(file_path):
        reg_path = r'Control Panel\Desktop'
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(reg_key, 'Wallpaper', 0, winreg.REG_SZ, file_path)
        winreg.SetValueEx(reg_key, 'WallpaperStyle', 0, winreg.REG_SZ, '10')
        winreg.CloseKey(reg_key)

    def refresh_desktop_wallpaper(self, image_file_with_path, image_desc):
        file_path = os.path.abspath(image_file_with_path)

        is_set = windll.user32.SystemParametersInfoW(self.SPI_SETDESKWALLPAPER, 0, file_path,
                                                     self.SPIF_UPDATEINIFILE | self.SPIF_SENDWININICHANGE)
        if is_set:
            print('Wallpaper updated. Enjoy!!!')
            balloon_tip('Bing', 'Wallpaper updated! \n {0}'.format(image_desc))
        else:
            print('Wallpaper could not be updated. Try again.')
            balloon_tip('Bing', 'Wallpaper could not be updated. Try again.')

    def download_image(self, servers, file_path, image_name):
        for server in servers:
            image_url = 'http:' + server + image_name
            try:
                urllib.request.urlretrieve(image_url, filename=file_path)
                break
            except urllib.error.HTTPError:
                continue  # we will try to download from all servers.

    def get_image_details(self, browse_data_url, image_details_url):
        browse_data = requests.get(browse_data_url).text
        data_regex = re.compile(".*browseData=({.*});.*")
        data_match = data_regex.match(browse_data)

        if data_match:
            data = data_match.group(1)
            json_data = json.loads(data)
            while True:
                # We select a random image id from the list of imageIds.
                image_id = random.choice(json_data['imageIds'])
                print(image_id)
                image_details = requests.get(image_details_url + image_id).text
                image_details_json = json.loads(image_details)
                image_name = image_details_json['wpFullFilename']
                if not image_name:
                    continue
                image_title = image_details_json['title']
                return image_name, image_title

        raise Exception('Could not connect to server. Please try again.')

if __name__ == '__main__':
    bing_random_images = BingRandomImages()
    bing_random_images.set_wallpaper()
