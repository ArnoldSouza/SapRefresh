# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import functools
from datetime import datetime
import time

from halo import Halo  # spinners for long running methods in terminal


def timeit(func):
    """Print the runtime of the decorated function"""
    @functools.wraps(func)
    def wrapper_timer(*args, **kwargs):
        start_time = time.perf_counter()  # 1
        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print('\n', f'Running {func.__name__!r} at: ', now)

        # start waiting spinner
        spinner = Halo(text='Loading', spinner='dots')
        spinner.start()

        value = func(*args, **kwargs)

        # stop waiting spinner
        spinner.stop()

        end_time = time.perf_counter()  # 2
        run_time = end_time - start_time  # 3
        if run_time < 60:
            print(f"Finished {func.__name__!r} in {run_time:.2f} seconds\n")
        if 60 <= run_time < 3600:
            run_time = run_time / 60  # converts to minutes
            print(f"Finished {func.__name__!r} in {run_time:.1f} minutes\n")
        if run_time >= 3600:
            run_time = run_time / 3600  # converts to minutes
            print(f"Finished {func.__name__!r} in {run_time:.1f} hours\n")
        return value
    return wrapper_timer


class SpinnerCursor(object):
    def __init__(self, text, spinner):
        """Create a spinner to show execution while waiting for processes"""
        self.spinner = Halo(text=text, spinner=spinner)

    def start(self):
        """start the animation of cursor"""
        self.spinner.start()

    def stop(self):
        """stop the animation of cursor"""
        self.spinner.succeed('End!')
