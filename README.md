# Java-poiexcel-demo
利用Apache POI组件读取U盘里的Excel文件内容并将其显示在Swing界面上  
**注意：已对代码进行简化处理，但不影响阅读思路**

该项目由一个main主方法和三个类构成，分别是ResFile、ConsumerUSBRoot、ProducerUSBRoot。
## main主方法
主方法调用File类的listRoots()方法来获取当前所有盘符然后记录下现有设备数量，然后实例化ResFile类并传入现有设备数量作为参数。ResFile类调用initUI()方法创建界面，同时创建两个线程t1和t2，分别创建一个生产者线程和一个消费者线程。最后启动线程监听系统是否有新增设备并进行相应的操作。
## ResFile
ResFile类继承JFrame，有initUI()方法创建界面，getAllFiles()方法获取所有文件名并显示在界面上，searchFile()方法查找资源（生产者），readFile()方法消费资源（消费者）。
## ConsumerUSBRoot
消费者类，实现Running接口并重写run方法，调用ResFile中的readFile方法来读取U盘里面的内容，当U盘拔出后通知生产者类继续判断是否有新设备插入。
## ProducerUSBRoot
生产者类，实现Runnable接口并重写run方法，调用ResFile中的searchFile方法来判断是否有新设备插入，如果没有则等待，如果有则通知消费者类。