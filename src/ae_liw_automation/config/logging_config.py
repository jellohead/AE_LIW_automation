# src/AE_311_automation/logging_config.py

import logging.config

LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,

    'formatters': {
        'standard': {
            'format': '[%(asctime)s] %(levelname)s in %(name)s: %(message)s'
        },
    },

    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'standard',
            'level': 'ERROR', # only show ERROR and above in console
        },
        'file': {
            'class': 'logging.FileHandler',
            'filename': 'automation.log',
            'formatter': 'standard',
            'level': 'DEBUG',
        },
    },

    'root': {
        # logging output sent to console and logfile
        'handlers': ['console', 'file'],
        'level': 'DEBUG',
    },

    'loggers': {
        'pptx': {
            'level': 'WARNING',
            'handlers': ['file'],
            'propagate': False,
        },
        'openpyxl': {
            'level': 'WARNING',
            'handlers': ['file'],
            'propagate': False,
        },
    },
}


def setup_logging():
    logging.config.dictConfig(LOGGING_CONFIG)
