# FVSData
This is an open source project, so please contribute or tell us your suggestions! The purpose of this project is to make it easier to extract data from different file types (currently supports csv and xlsx files). It was originally made for the NJ State Department of Forestry, to make it easier to deal with simulation programs that require these files as inputs. This tool helps the foresters work more efficiently, and we are currently adding more functionalities to it. 

The runner.py is the prototype, with basic functionalities. It can take in xlsx or csv files, and then will print out all the columns in the file. You can then type in which columns you want, and whether you want these columns and data to be exported as a csv or an xlsx.

The FVSData.py is the main library. You can use the test.py to see how it works, but essentially you can make an object of type FVSData, and call the methods to do the things you need to do. On top of adding much more functionality a webframework is in the works so that you can just upload the files directly to us, do the manipulations you need to do, and then can download the resulting file!
