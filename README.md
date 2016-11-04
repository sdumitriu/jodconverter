# XWiki JODConverter

The XWiki project has forked the JODConverter project because it seems dead and we need to bring some fixes to it. We hope that someone from the JODConverter project will be able to apply our fixes at some point.

## JODConverter

JODConverter (Java OpenDocument Converter) automates document conversions using LibreOffice or OpenOffice.org.

See the [Google Code](http://code.google.com/p/jodconverter/) project for more info.

## Build

In order to build the project you need to execute:

    mvn clean install -DskipTests

If you want to run the unit tests you need to:

* Download the [libsigar](https://github.com/hyperic/sigar) binary corresponding to your OS (e.g. ``libsigar-amd64-linux.so``)
* Create the ``jodconverter-core/lib`` folder and put the libsigar binary there
* ``mvn clean install -Djava.library.path=lib``
* Note that currently the conversion to/from ``sxw`` format is failing

        Caused by: com.sun.star.task.ErrorCodeIOException: SfxBaseModel::impl_store <file:///tmp/test311448721067585328.sxw> failed: 0x81a
