import logging
from PyQt6.QtCore import QDateTime

current_datetime = QDateTime.currentDateTime()
#datetime_str = current_datetime.toString("hh_mm_ss_dd_MM_yyyy")
datetime_str = current_datetime.toString("dd_MM_yyyy")

# Set up logging to only file
#logging.basicConfig(level=logging.INFO,
 #                   format='%(asctime)s [%(levelname)s]: %(message)s',
  #                  handlers=[logging.FileHandler(f'Logs/Session_{datetime_str}.log')])


# Set up logging to both file and console
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s [%(levelname)s]: %(message)s',
                    handlers=[
                        logging.FileHandler(f'Logs/Session_{datetime_str}.log'),
                        logging.StreamHandler()  # This sends logs to console
                    ])

logger = logging.getLogger('my_app')
