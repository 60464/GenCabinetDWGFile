# 机柜图纸自动生成工具 GenCabinetDWGFile  
Cabinet drawings are automatically generated in AUTOCAD format based on the configuration file

# 目的  
计算机联锁系统机柜生产图纸的绘制，是联锁项目实施过程中的重要工作。不同的项目，其机柜配置和各种板卡配置的差异很大，由于系统本身板卡类型和配置组合过多，导致从产品化图纸到项目图纸的转化过程需要花费更多的人工。既有的那种按照车站大、中、小规模制作产品化图纸的方式已经无法满足新项目的需求。基于这样的现状，为了提高工程师的工作效率和准确率，开发了联锁项目机柜生产图纸自动生成工具。  
The drawing of cabinet production drawing of computer interlocking system is an important work in the process of interlocking project implementation. For different projects, the cabinet configuration and various board configurations vary greatly. Due to the excessive combination of board types and configurations of the system itself, the conversion process from production drawings to project drawings requires more labor. The existing way of producing production drawings according to the large, medium and small size of the station can no longer meet the needs of the new project. Based on this situation, in order to improve the work efficiency and accuracy of engineers, the automatic generation tool of interlocking project cabinet production drawings is developed.  

# 说明  
软件采用EXECL配置文件的模式对项目的机柜需求进行简单配置，而且EXECL配置文件中的设备名称、高度等信息不需要输入，可以直接通过下拉菜单选择。同时可以输入对每个设备的备注信息，即在机柜图纸中每个设备旁边进行文字说明。  
The software uses the EXECL configuration file to simply configure the cabinet requirements of a project. In addition, the device name and height in the EXECL configuration file can be directly selected from the drop-down menu. In addition, you can enter remarks for each device, that is, write descriptions next to each device in the cabinet drawing. 
将各种设备的标准图纸放在DWG文件夹中，软件生成机柜图纸时会自动调用DWG文件夹中的标准图纸。  
The standard drawings of various devices are placed in the DWG folder, and the standard drawings in the DWG folder will be automatically invoked when the software generates the cabinet drawings.  


