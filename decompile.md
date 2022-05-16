https://github.com/mrquincle/roveropen/blob/master/src/org/almende/roveropen/RoverOpenActivity.java
github上的开源方案，可以参考

https://github.com/darwinbeing/RaceTrackRover
python方案

https://manuals.brookstone.com/792593p_manual.pdf
说明书

https://github.com/yocontra/rover
js开源方案

https://github.com/simondlevy/RoverPylot
ps3的遥控器，通过pc或者笔记本，操作小车

http://www.wifi-robots.com/forum.php?mod=viewthread&tid=5476&extra=&highlight=psp&page=1
psp的遥控方案

https://zhuanlan.zhihu.com/p/19896812
知乎的文章，数据是加密的

拍照时执行
AppCommand.getAppCommandInstace().sendCommand(8); 
AppCameraShootingFunction.getAppCameraShootingFunctionInstance().ShootingInit((AppThread.getAppThreadInstance()).CurrentVideoType);

this.Stealth_Btn) 隐身模式
if (!AppInforToCustom.getAppInforToCustomInstance().getIsSt	ealthControl()) {
AppNightLightFunction.getAppNightLightFunctionInstance().NightLight_OandC(Boolean.valueOf(true));


//创建连接框
              case 1002:
                WificarNew.this.connectionProgressDialog1.cancel();
                Connect_Dialog.createconnectDialog((Context)WificarNew.instance).show();
//连接进度条
              case 1004:
                WificarNew.this.connectionProgressDialog1.show();
//进度条消失
              case 1003:
                WificarNew.this.connectionProgressDialog1.cancel();
                WificarNew.this.CheckFirstApp();


//初始化socket
boolean bool = AppConnect.getInstance((Context)WificarNew.instance, WificarNew.this.bigeyeCallBack).initSocket();


反编译工具，jad，jd-gui，dex2jar，DJ Java Decompiler 3.8
