import re
import os
import random
import json
import requests
import urllib.request
import winreg
from tray_notify import balloon_tip
if os.name is 'nt':
    from ctypes import windll
else:
    raise Exception('Unsupported OS.')


class BingWallpaper:
    URL = r'http://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=99'
    SPIF_UPDATEINIFILE = 0x01
    SPIF_SENDWININICHANGE = 0x02
    SPI_SETDESKWALLPAPER = 0x0014

    def __init__(self):
        self._daily_wall_file = 'wall_of_day.jpg'
        self._image_desc = ''

    def set_wallpaper(self):
        try:
            image_data = json.loads(requests.get(self.URL).text)

            valid_images = [image for image in image_data['images'] if image['wp']]
            total_images = len(valid_images)
            random_image_id = random.randint(0, total_images - 1)
            selected_image = valid_images[random_image_id]

            image_url = 'http://www.bing.com' + selected_image['url']
            self._image_desc = selected_image['copyright']

            # url for better quality image
            image_download_url = 'http://www.bing.com/hpwp/' + selected_image['hsh']
            image_name = image_url[re.search("rb/", image_url).end():re.search('_EN', image_url).start()] + '.jpg'
            print(image_name)

            bing_dir = r'C:\Users\{0}\Pictures\Bing'.format(os.getlogin())
            if not os.path.isdir(bing_dir):
                os.mkdir(bing_dir)

            file_path = r'{0}\{1}'.format(bing_dir, self._daily_wall_file)
            if os.path.exists(file_path):
                os.remove(file_path)

            self._download_image(file_path, image_download_url, image_url)

            self._update_registry_values(file_path)
            self._refresh_desktop_wallpaper(file_path)
        except Exception as e:
            print('Wallpaper can\'t be updated!')
            print(e)

    @staticmethod
    def _download_image(file_path, image_download_url, image_url):
        try:
            # try downloading image from hash url
            urllib.request.urlretrieve(image_download_url, filename=file_path)
        except urllib.error.HTTPError as e:
            # try downloading from default url
            urllib.request.urlretrieve(image_url, filename=file_path)
            print(e)

    @staticmethod
    def _update_registry_values(file_path):
        reg_path = r'Control Panel\Desktop'
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(reg_key, 'Wallpaper', 0, winreg.REG_SZ, file_path)
        winreg.SetValueEx(reg_key, 'WallpaperStyle', 0, winreg.REG_SZ, '10')
        winreg.CloseKey(reg_key)

    def _refresh_desktop_wallpaper(self, image_file_with_path):
        file_path = os.path.abspath(image_file_with_path)

        is_set = windll.user32.SystemParametersInfoW(self.SPI_SETDESKWALLPAPER, 0, file_path,
                                                     self.SPIF_UPDATEINIFILE | self.SPIF_SENDWININICHANGE)
        if is_set:
            print('Wallpaper updated. Enjoy!!!')
            balloon_tip('Bing', 'Wallpaper updated! \n {0}'.format(self._image_desc))
        else:
            print('Wallpaper could not be updated. Try again.')
            balloon_tip('Bing', 'Wallpaper could not be updated. Try again.')


if __name__ == '__main__':
    bing = BingWallpaper()
    bing.set_wallpaper()
