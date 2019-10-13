multiply-channels
===

This project provides a commandline tool to create knxprod files for ETS.

It is motivated by this [knx](https://github.com/thelsing/knx) stack and the according [CreateKnxProd](https://github.com/thelsing/CreateKnxProd) tool.
A first (initial) xml file should be created with [CreateKnxProd](https://github.com/thelsing/CreateKnxProd), there is currently no new command available in multiply-channels.

The current version provides the following verbs:

- **create**     Check given xml file and create knxprod

- **check**      Execute sanity checks on given xml file

- **knxprod**    Create knxprod file from given xml file without checks

- **help**       Display more information on a specific command.

- **version**    Display version information.

This tool also creates a header file with defines for all parameter and com object definitions in the xml. An ETS must be installed on the PC. Currently ETS 4, ETS 5.5, ETS 5.6 and ETS 5.7 are supported. The correct ETS converter version is automatically found dependent on xmlns of provided xml document.
This project uses dotnet core 3.0.

### Examples:

- multiplychannels create Sensor

    Reads Sensor.xml, do sanity chacks, produce Sensor.h, produce Sensor.knxprod

- multiplychannels knxprod -o device.knxpord Sensor.xml

    Reads Sensor.xml, produce device.knxprod

- multiplychannels check Sensor.xml

    Reads Sensor.xml, do sanity checks
