Python Service Launcher for Windows (Python Cron) v0.9
======================================================
Python Tool for Windows Service Loading and Scheduling

Usage Example - LogWriterSvc.py
-------------------------------
This script opens a log and writes to it with ``serviceCron.tab`` scheduling:

.. code-block:: python

    from AbstractServiceClass import AbstractServiceClass

    class LogWriter(AbstractServiceClass):
        def __init__(self):
            AbstractServiceClass.__init__(self)

        def __doit__(self):

            """ Put your code in Here Plz :) """

            self.log('Too much work in here...')

Format Rules to Follow
----------------------
- You'll have to use ``LogWriterSvc.py`` as a template for your custom scripts.
- Script name must be ``<class name>Svc.py`` for ``ServiceLauncher`` to load it as such thing, and it must be placed in ``services\`` folder. All other files into ``services\`` folder are to be consumed by services themselves or ``AbstractServiceClass``.
- Every script must have its own cron rule into ``services\serviceCron.tab``. This rules follow usual Unix cron format.

Prerequisites
-------------
Didn't I say this doesn't work on its own?... No I didn't, I know.

Ok, so for this to work you'll probably need to download some packages, of course. Less usual are ``pywin32``, ``python-crontab`` (not ``crontab``), ``croniter``.

How to enjoy
------------
Just add your favourite scripts with the given format to ``services\`` folder, incorporate a cron rule for each one of them, run ``python <class name>Svc.py install`` for Windows to load main service, grab some popcorn, dim the lights and do a final ``python <class name>Svc.py start`` to start the service.

## Have Fun!

P.d.: Project incorporates a basic ``setup.py`` file for py2exe. It doesn't grab your scripts to make the executable file, so give it a try just in case you had any dependency problem (Not extensively tested!).