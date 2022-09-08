package org.example;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellReference;
import sun.awt.shell.ShellFolder;

import javax.swing.*;
import java.awt.*;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/*
 * 实现当检测到有设备插入时，查找设备中的所有文件，使用生产者消费者模式进行检查和查找
 * 1.生产者：在设备插入之前先判断系统开始的盘符数，
 * 然后创建一个线程不断判断系统有多少个盘符，若判断出盘符数增多，则该线程等待并唤醒消费者；否则一直判断。
 * 2.消费者：在没有判断出有插入设备时，处于等待状态；若有，则查找设备中是否包含指定文件。
 * 3.资源：将插入的设备当作资源
 */
public class App {
    public static void main(String[] args) {
        final File[] dir = File.listRoots(); //获取本机所有盘符
        final int count = dir.length; //现有设备数
        ResFile rf = new ResFile(count);
        rf.initUI();
        Thread t1 = new Thread(new ProducerUSBRoot(rf));
        Thread t2 = new Thread(new ConsumerUSBRoot(rf));
        t1.start();
        t2.start();
    }

    //资源
    public static class ResFile extends JFrame {
        private int count = 0; //初始化计数器
        private boolean flag = false; //判断是否有设备插入的标记
        private File[] dirs;
        private File[] files; //获取所有文件名
        private int i; //public synchronized void readFile()的for循环

        //设置一个全局动态数组，来存放文件路径
        //主要遍历文件夹，包含所有子文件夹，文件的情况时，用到递归，所以要这样设置
        public static ArrayList<File> dirAllStrArr = new ArrayList<>();

        public ResFile(int count) {
            this.count = count;
        }

        final JTextField textFieldType = new JTextField(50); //设置型号文本框
        final JTextField textFieldNumber = new JTextField(50); //设置序列号文本框
        final JTextField textFieldTunnalMin = new JTextField(50); //设置通道最小值提示文本框
        final JTextField textFieldTunnalMax = new JTextField(50); //设置通道最大值提示文本框
        final JTextField textFieldValueStability = new JTextField(50); //设置示值稳定性提示文本框
        final JTextField textFieldValueTolerance = new JTextField(50); //设置吸光度示值误差提示文本框
        final JTextField textFieldAbsorbanceRepeat = new JTextField(50); //设置吸光度重复值提示文本框
        final JTextField textFieldChannelDifference = new JTextField(50); //设置通道差异提示文本框
        final JTextField textFieldSoftwareFolder = new JTextField(50); //设置软件提示文本框
        final JTextField textFieldSystemGHO = new JTextField(50); //设置系统文件提示文本框
        Font fontLabel = new Font("宋体", Font.PLAIN, 15);//设置标签字体
        Font fontNumber = new Font("宋体", Font.BOLD, 60);//设置型号和序列号字体
        Font fontResult = new Font("宋体", Font.PLAIN, 14);//设置结果字体格式
        final JLabel labelTunnalMinResult = new JLabel(); //设置通道最小值结果显示标签字样
        final JLabel labelTunnalMaxResult = new JLabel(); //设置通道最大值结果显示标签字样
        final JLabel labelValueStabilityResult = new JLabel(); //设置示值稳定性结果显示标签字样
        final JLabel labelValueToleranceResult = new JLabel(); //设置吸光度示值误差结果显示标签字样
        final JLabel labelAbsorbanceRepeatResult = new JLabel(); //设置吸光度重复值结果显示标签字样
        final JLabel labelChannelDifferenceResult = new JLabel(); //设置通道差异结果显示标签字样
        final JLabel labelSoftwareFolderResult = new JLabel(); //设置软件显示标签字样
        final JLabel labelSystemGHOResult = new JLabel(); //设置系统文件显示标签字样

        //创建界面
        public void initUI() {
            JFrame jFrame = new JFrame();
            this.setTitle("报告信息提取"); //设置标题
            this.setSize(500, 470);//设置大小
            this.setDefaultCloseOperation(EXIT_ON_CLOSE);
            this.setResizable(false); //设置不可拉伸
            this.setLayout(null); //关闭流式布局

            //开始界面元素设计
            /*---------------型号文本框---------------*/
            textFieldType.setBounds(25, 10, 435, 55);//设置位置大小
            textFieldType.setEditable(false);//设置不可编辑
            textFieldType.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldType);
            /*---------------型号文本框---------------*/
            /*---------------序列号文本框---------------*/
            textFieldNumber.setBounds(25, 71, 435, 55); //设置位置大小
            textFieldNumber.setEditable(false);//设置不可编辑
            textFieldNumber.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldNumber); //添加标签
            /*---------------序列号文本框---------------*/
            /*---------------示值稳定性标签---------------*/
            final JLabel labelValueStability = new JLabel("示值稳定性（±0.002）："); //设置目录字样
            labelValueStability.setBounds(25, 136, 255, 25); //设置位置大小
            labelValueStability.setFont(fontLabel);
            this.add(labelValueStability); //添加标签
            /*---------------示值稳定性标签---------------*/
            /*---------------示值稳定性文本框---------------*/
            textFieldValueStability.setBounds(235, 133, 150, 30); //设置位置大小
            textFieldValueStability.setEditable(false);//设置不可编辑
            textFieldValueStability.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldValueStability); //添加标签
            /*---------------示值稳定性文本框---------------*/
            /*---------------吸光度示值误差标签---------------*/
            final JLabel labelValueTolerance = new JLabel("吸光度示值误差（±0.015）："); //设置目录字样
            labelValueTolerance.setBounds(25, 170, 255, 25); //设置位置大小
            labelValueTolerance.setFont(fontLabel);
            this.add(labelValueTolerance); //添加标签
            /*---------------吸光度示值误差标签---------------*/
            /*---------------吸光度示值误差文本框---------------*/
            textFieldValueTolerance.setBounds(235, 168, 150, 30); //设置位置大小
            textFieldValueTolerance.setEditable(false);//设置不可编辑
            textFieldValueTolerance.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldValueTolerance); //添加标签
            /*---------------吸光度示值误差文本框---------------*/
            /*---------------吸光度重复值标签---------------*/
            final JLabel labelAbsorbanceRepeat = new JLabel("吸光度重复值（<0.1）："); //设置目录字样
            labelAbsorbanceRepeat.setBounds(25, 206, 255, 25); //设置位置大小
            labelAbsorbanceRepeat.setFont(fontLabel);
            this.add(labelAbsorbanceRepeat); //添加标签
            /*---------------吸光度重复值标签---------------*/
            /*---------------吸光度重复值文本框---------------*/
            textFieldAbsorbanceRepeat.setBounds(235, 202, 150, 30); //设置位置大小
            textFieldAbsorbanceRepeat.setEditable(false);//设置不可编辑
            textFieldAbsorbanceRepeat.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldAbsorbanceRepeat); //添加标签
            /*---------------吸光度重复值文本框---------------*/
            /*---------------通道差异标签---------------*/
            final JLabel labelChannelDifference = new JLabel("通道差异（≤0.02）："); //设置目录字样
            labelChannelDifference.setBounds(25, 276, 255, 25); //设置位置大小
            labelChannelDifference.setFont(fontLabel);
            this.add(labelChannelDifference); //添加标签
            /*---------------通道差异标签---------------*/
            /*---------------通道差异文本框---------------*/
            textFieldChannelDifference.setBounds(235, 272, 150, 30); //设置位置大小
            textFieldChannelDifference.setEditable(false);//设置不可编辑
            textFieldChannelDifference.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldChannelDifference); //添加标签
            /*---------------通道差异文本框---------------*/
            /*---------------设置示值稳定性结果显示标签字样---------------*/
            labelValueStabilityResult.setBounds(395, 136, 30, 25); //设置位置大小
            labelValueStabilityResult.setFont(fontLabel);//设置字体
            //labelValueStabilityResult.setText("正确");//测试显示效果
            this.add(labelValueStabilityResult);//添加标签
            /*---------------设置示值稳定性结果显示标签字样---------------*/
            /*---------------设置吸光度示值误差结果显示标签字样---------------*/
            labelValueToleranceResult.setBounds(395, 170, 30, 25); //设置位置大小
            labelValueToleranceResult.setFont(fontLabel);//设置字体
            //labelValueToleranceResult.setText("正确");//测试显示效果
            this.add(labelValueToleranceResult);//添加标签
            /*---------------设置吸光度示值误差结果显示标签字样---------------*/
            /*---------------设置吸光度重复值结果显示标签字样---------------*/
            labelAbsorbanceRepeatResult.setBounds(395, 206, 30, 25); //设置位置大小
            labelAbsorbanceRepeatResult.setFont(fontLabel);//设置字体
            //labelAbsorbanceRepeatResult.setText("正确");//测试显示效果
            this.add(labelAbsorbanceRepeatResult);//添加标签
            /*---------------设置吸光度重复值结果显示标签字样---------------*/
            /*---------------设置通道差异结果显示标签字样---------------*/
            labelChannelDifferenceResult.setBounds(395, 276, 30, 25); //设置位置大小
            labelChannelDifferenceResult.setFont(fontLabel);//设置字体
            //labelChannelDifferenceResult.setText("正确");//测试显示效果
            this.add(labelChannelDifferenceResult);//添加标签
            /*---------------设置通道差异结果显示标签字样---------------*/

            /*---------------设置分割线标签字样---------------*/
            final JLabel labelSplitLine = new JLabel("------------------------------------------------------------------------------------------------------------"); //这是分割线标签
            labelSplitLine.setBounds(25, 306, 460, 25); //设置位置大小
            this.add(labelSplitLine); //添加标签
            /*---------------设置分割线标签字样---------------*/

            /*---------------设置软件标签---------------*/
            final JLabel labelSoftwareFolder = new JLabel("软件文件夹："); //设置软件文件夹标签字样
            labelSoftwareFolder.setBounds(25, 345, 255, 25); //设置位置大小
            labelSoftwareFolder.setFont(fontLabel); //设置字体
            this.add(labelSoftwareFolder); //添加标签
            /*---------------设置软件标签---------------*/
            /*---------------设置软件显示文本框---------------*/
            textFieldSoftwareFolder.setBounds(180, 340, 205, 30); //设置位置大小
            textFieldSoftwareFolder.setEditable(false); //设置不可编辑
            textFieldSoftwareFolder.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldSoftwareFolder); //添加标签
            /*---------------设置软件显示文本框---------------*/
            /*---------------设置软件文件夹显示标签字样---------------*/
            labelSoftwareFolderResult.setBounds(395, 335, 90, 35);
            //labelSoftwareFolderResult.setFont(fontLabel); //设置字体
//            labelSoftwareFolderResult.setText("测试");//测试显示效果
            this.add(labelSoftwareFolderResult);//添加标签
            /*---------------设置软件文件夹显示标签字样---------------*/

            /*---------------设置系统文件标签---------------*/
            final JLabel labelSystemGHO = new JLabel("系统文件：");//设置系统文件标签字样
            labelSystemGHO.setBounds(25, 380, 255, 25); //设置位置大小
            labelSystemGHO.setFont(fontLabel); //设置字体
            this.add(labelSystemGHO); //添加标签
            /*---------------设置系统文件标签---------------*/
            /*---------------设置系统文件显示文本框---------------*/
            textFieldSystemGHO.setBounds(180, 375, 205, 30); //设置位置大小
            textFieldSystemGHO.setEditable(false); //设置不可编辑
            textFieldSystemGHO.setBackground(Color.white); //设置文本框背景色为白色
            this.add(textFieldSystemGHO);//添加文本框
            /*---------------设置系统文件显示文本框---------------*/
            /*---------------设置系统文件显示标签字样---------------*/
            labelSystemGHOResult.setBounds(395, 377, 60, 25);//设置位置大小
            labelSystemGHOResult.setFont(fontLabel);
            //labelSystemGHOResult.setText("测试");//测试显示效果
            this.add(labelSystemGHOResult);
            /*---------------设置系统文件显示标签字样---------------*/
            //结束界面元素设计

            this.setVisible(true); //设置窗体可见
        }
        //创建界面

        /*-----遍历文件夹文件，存在放全局动态数组中-----*/
        public static ArrayList<File> FindFile(File dir) throws Exception {
            final File[] files = dir.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        FindFile(file);
                    } else {
                        dirAllStrArr.add(file);
                        if (file.getName().equals("App.exe")) {
                            dirAllStrArr.set(0, file);
                        }
                        if ((file.getName().contains(".gho")) || (file.getName().contains(".GHO"))) {
                            dirAllStrArr.set(1, file);
                        }
                        if (file.getName().contains(".xls")) {
                            String pattern = ".*FQC.*";
                            String pattern1 = ".*Microplate.*";
                            String pattern2 = ".*酶标.*";
                            boolean isMatch1 = Pattern.matches(pattern, file.toString());
                            boolean isMatch2 = Pattern.matches(pattern1, file.toString());
                            boolean isMatch3 = Pattern.matches(pattern2, file.toString());
                            if (isMatch1 || isMatch2 || isMatch3) {
                                dirAllStrArr.set(2, file);
                            }
                        }
                    }
                }
            }
            return dirAllStrArr;
        }
        /*-----遍历文件夹文件，存在放全局动态数组中-----*/

        /*-----获取所有文件名-----*/
        public void getAllFiles(File dir) throws Exception {
            dirAllStrArr.clear();//清空数组内容
            FindFile(dir);//调用查找文件方法
            /*-----方法2-----*/
            if (dirAllStrArr.size() <= 2) {
                textFieldSoftwareFolder.setText("");
                labelSoftwareFolderResult.setIcon(null);
                textFieldSystemGHO.setText("");
                textFieldType.setText("");
                textFieldNumber.setText("");
                textFieldValueStability.setText("");
                textFieldValueTolerance.setText("");
                textFieldAbsorbanceRepeat.setText("");
                textFieldChannelDifference.setText("");
                labelTunnalMinResult.setText("");
                labelTunnalMaxResult.setText("");
                labelValueStabilityResult.setText("");
                labelValueToleranceResult.setText("");
                labelAbsorbanceRepeatResult.setText("");
                labelChannelDifferenceResult.setText("");
                return;
            }

            if (dirAllStrArr.get(0).getName().contains(".exe")) {
                textFieldSoftwareFolder.setText(dirAllStrArr.get(0).getAbsolutePath().toString());
                final File filePath = new File(dirAllStrArr.get(0).getAbsolutePath());
                final ShellFolder shellFolder = ShellFolder.getShellFolder(filePath); //获取EXE路径
                final ImageIcon imageIcon = new ImageIcon(shellFolder.getIcon(true)); //获取图标
                imageIcon.getImage().flush(); //刷新旧图标（解决缓存问题）
                labelSoftwareFolderResult.setIcon(imageIcon); //标签显示新图标
            } else {
                textFieldSoftwareFolder.setText("");
                labelSoftwareFolderResult.setIcon(null);
            }

            if (dirAllStrArr.get(1).getName().contains(".gho") || dirAllStrArr.get(1).getName().contains(".GHO")) {
                textFieldSystemGHO.setText(dirAllStrArr.get(1).getAbsolutePath().toString());
            } else {
                textFieldSystemGHO.setText("");
            }

            if (dirAllStrArr.get(2).getName().contains(".xls")) {

                //System.out.println("file: " + f); //打印完整路径
                //System.out.println(f.getName()+"\n===================="); //打印文件名
                /*----------读取U盘XLS文件的数据----------*/
                POIFSFileSystem fileSystem = null;
                HSSFWorkbook workbook = null;
                try {
                    fileSystem = new POIFSFileSystem(new BufferedInputStream(new FileInputStream(dirAllStrArr.get(2).getAbsoluteFile()))); //使用BufferedInputStream读取
                    workbook = new HSSFWorkbook(fileSystem);
                } catch (IOException e) {
                    e.printStackTrace();
                }
                HSSFSheet sheet = workbook.getSheetAt(0);
                /*-----判断中英文报告-----*/
                final CellReference cellReferenceA1 = new CellReference("A1");
                final HSSFRow rowA1 = sheet.getRow(cellReferenceA1.getRow());
                final HSSFCell cellA1 = rowA1.getCell(cellReferenceA1.getCol());
                final String enReport = cellA1.getStringCellValue();
                /*-----判断中英文报告-----*/
                /*-----英文报告-----*/
                if (enReport.equals("Microplate reader Quality performance report")) {
                    //......
                }
                /*-----英文报告-----*/
                /*-----中文报告-----*/
                if (!enReport.equals("Microplate reader Quality performance report")) {
                    /*-----中文报告-----*/
                    /*-----读取型号单元格数据-----*/
                    Pattern pattern = Pattern.compile("[A-Z](.*?)\\d$");
                    final CellReference cellReferenceP4 = new CellReference("P4");
                    final HSSFRow rowP4 = sheet.getRow(cellReferenceP4.getRow());
                    final HSSFCell cellP4 = rowP4.getCell(cellReferenceP4.getCol());
                    String cnProductType = cellP4.getStringCellValue();
                    Matcher cnmatcherProductType = pattern.matcher(cnProductType.replaceAll("\\s", ""));
                    /*-----读取型号单元格数据-----*/
                    /*-----读取序列号单元格数据-----*/
                    final CellReference cellReferenceP5 = new CellReference("P5");
                    final HSSFRow rowP5 = sheet.getRow(cellReferenceP5.getRow());
                    final HSSFCell cellP5 = rowP5.getCell(cellReferenceP5.getCol());
                    String cnProductNumber = cellP5.getStringCellValue();
                    final Matcher cnmatcherProductNumber = pattern.matcher(cnProductNumber.replaceAll("\\s", ""));
                    /*-----读取序列号单元格数据-----*/
                    if (cnmatcherProductType.group().equals("Product1")) {
                        textFieldType.setFont(fontNumber);
                        textFieldType.setText(cnmatcherProductType.group());
                        textFieldNumber.setFont(fontNumber);
                        textFieldNumber.setText(cnmatcherProductNumber.group());
                        /*-----读取示值稳定性单元格-----*/
                        final CellReference cellReferenceF38 = new CellReference("F38");//5min
                        final HSSFRow rowF38 = sheet.getRow(cellReferenceF38.getRow());
                        final HSSFCell cellF38 = rowF38.getCell(cellReferenceF38.getCol());
                        final CellReference cellReferenceK38 = new CellReference("K38");//10min
                        final HSSFRow rowK38 = sheet.getRow(cellReferenceK38.getRow());
                        final HSSFCell cellK38 = rowK38.getCell(cellReferenceK38.getCol());
                        final CellReference cellReferenceA38 = new CellReference("A38");//初值
                        final HSSFRow rowA38 = sheet.getRow(cellReferenceA38.getRow());
                        final HSSFCell cellA38 = rowA38.getCell(cellReferenceA38.getCol());
                        final double resultA38 = (Math.max(cellF38.getNumericCellValue(), cellK38.getNumericCellValue()))
                                - cellA38.getNumericCellValue();//最大减去最小值
                        String resultA38String = String.format("%.4f", resultA38);
                        textFieldValueStability.setFont(fontResult);//设置字体
                        textFieldValueStability.setText(resultA38String);
                        if (resultA38 >= 0.002 || resultA38 <= -0.002) {
                            labelValueStabilityResult.setText("错误");
                            labelValueStabilityResult.setForeground(Color.red);
                        } else {
                            labelValueStabilityResult.setText("正确");
                            labelValueStabilityResult.setForeground(Color.green);
                        }
                        /*-----读取示值稳定性单元格-----*/
                        /*-----读取吸光度单元格-----*/
                        /*-----------------405-----------------*/
                        final CellReference cellReferenceD44 = new CellReference("D44");//0.2标准值
                        final HSSFRow rowD44 = sheet.getRow(cellReferenceD44.getRow());
                        final HSSFCell cellD44 = rowD44.getCell(cellReferenceD44.getCol());
                        final CellReference cellReferenceG44 = new CellReference("G44");
                        final HSSFRow rowG44 = sheet.getRow(cellReferenceG44.getRow());
                        final HSSFCell cellG44 = rowG44.getCell(cellReferenceG44.getCol());
                        final CellReference cellReferenceJ44 = new CellReference("J44");
                        final HSSFRow rowJ44 = sheet.getRow(cellReferenceJ44.getRow());
                        final HSSFCell cellJ44 = rowJ44.getCell(cellReferenceJ44.getCol());
                        final CellReference cellReferenceM44 = new CellReference("M44");
                        final HSSFRow rowM44 = sheet.getRow(cellReferenceM44.getRow());
                        final HSSFCell cellM44 = rowM44.getCell(cellReferenceM44.getCol());
                        final double resultS44 = ((
                                cellG44.getNumericCellValue() + cellJ44.getNumericCellValue() + cellM44.getNumericCellValue()) / 3)
                                - cellD44.getNumericCellValue();
                        final CellReference cellReferenceD45 = new CellReference("D45");
                        final HSSFRow rowD45 = sheet.getRow(cellReferenceD45.getRow());
                        final HSSFCell cellD45 = rowD45.getCell(cellReferenceD45.getCol());
                        final CellReference cellReferenceG45 = new CellReference("G45");
                        final HSSFRow rowG45 = sheet.getRow(cellReferenceG45.getRow());
                        final HSSFCell cellG45 = rowG45.getCell(cellReferenceG45.getCol());
                        final CellReference cellReferenceJ45 = new CellReference("J45");
                        final HSSFRow rowJ45 = sheet.getRow(cellReferenceJ45.getRow());
                        final HSSFCell cellJ45 = rowJ45.getCell(cellReferenceJ45.getCol());
                        final CellReference cellReferenceM45 = new CellReference("M45");
                        final HSSFRow rowM45 = sheet.getRow(cellReferenceM45.getRow());
                        final HSSFCell cellM45 = rowM45.getCell(cellReferenceM45.getCol());
                        final double resultS45 = ((
                                cellG45.getNumericCellValue() + cellJ45.getNumericCellValue() + cellM45.getNumericCellValue()) / 3)
                                - cellD45.getNumericCellValue();
                        final CellReference cellReferenceD46 = new CellReference("D46");
                        final HSSFRow rowD46 = sheet.getRow(cellReferenceD46.getRow());
                        final HSSFCell cellD46 = rowD46.getCell(cellReferenceD46.getCol());
                        final CellReference cellReferenceG46 = new CellReference("G46");
                        final HSSFRow rowG46 = sheet.getRow(cellReferenceG46.getRow());
                        final HSSFCell cellG46 = rowG46.getCell(cellReferenceG46.getCol());
                        final CellReference cellReferenceJ46 = new CellReference("J46");
                        final HSSFRow rowJ46 = sheet.getRow(cellReferenceJ46.getRow());
                        final HSSFCell cellJ46 = rowJ46.getCell(cellReferenceJ46.getCol());
                        final CellReference cellReferenceM46 = new CellReference("M46");
                        final HSSFRow rowM46 = sheet.getRow(cellReferenceM46.getRow());
                        final HSSFCell cellM46 = rowM46.getCell(cellReferenceM46.getCol());
                        final double resultS46 = ((
                                cellG46.getNumericCellValue() + cellJ46.getNumericCellValue() + cellM46.getNumericCellValue()) / 3)
                                - cellD46.getNumericCellValue();
                        final CellReference cellReferenceD47 = new CellReference("D47");
                        final HSSFRow rowD47 = sheet.getRow(cellReferenceD47.getRow());
                        final HSSFCell cellD47 = rowD47.getCell(cellReferenceD47.getCol());
                        final CellReference cellReferenceG47 = new CellReference("G47");
                        final HSSFRow rowG47 = sheet.getRow(cellReferenceG47.getRow());
                        final HSSFCell cellG47 = rowG47.getCell(cellReferenceG47.getCol());
                        final CellReference cellReferenceJ47 = new CellReference("J47");
                        final HSSFRow rowJ47 = sheet.getRow(cellReferenceJ47.getRow());
                        final HSSFCell cellJ47 = rowJ47.getCell(cellReferenceJ47.getCol());
                        final CellReference cellReferenceM47 = new CellReference("M47");
                        final HSSFRow rowM47 = sheet.getRow(cellReferenceM47.getRow());
                        final HSSFCell cellM47 = rowM47.getCell(cellReferenceM47.getCol());
                        final double resultS47 = ((
                                cellG47.getNumericCellValue() + cellJ47.getNumericCellValue() + cellM47.getNumericCellValue()) / 3)
                                - cellD47.getNumericCellValue();
                        final CellReference cellReferenceD48 = new CellReference("D48");
                        final HSSFRow rowD48 = sheet.getRow(cellReferenceD48.getRow());
                        final HSSFCell cellD48 = rowD48.getCell(cellReferenceD48.getCol());
                        final CellReference cellReferenceG48 = new CellReference("G48");
                        final HSSFRow rowG48 = sheet.getRow(cellReferenceG48.getRow());
                        final HSSFCell cellG48 = rowG48.getCell(cellReferenceG48.getCol());
                        final CellReference cellReferenceJ48 = new CellReference("J48");
                        final HSSFRow rowJ48 = sheet.getRow(cellReferenceJ48.getRow());
                        final HSSFCell cellJ48 = rowJ48.getCell(cellReferenceJ48.getCol());
                        final CellReference cellReferenceM48 = new CellReference("M48");
                        final HSSFRow rowM48 = sheet.getRow(cellReferenceM48.getRow());
                        final HSSFCell cellM48 = rowM48.getCell(cellReferenceM48.getCol());
                        final double resultS48 = ((
                                cellG48.getNumericCellValue() + cellJ48.getNumericCellValue() + cellM48.getNumericCellValue()) / 3)
                                - cellD48.getNumericCellValue();
                        /*-----------------405-----------------*/
                        /*-----------------450-----------------*/
                        final CellReference cellReferenceD49 = new CellReference("D49");
                        final HSSFRow rowD49 = sheet.getRow(cellReferenceD49.getRow());
                        final HSSFCell cellD49 = rowD49.getCell(cellReferenceD49.getCol());
                        final CellReference cellReferenceG49 = new CellReference("G49");
                        final HSSFRow rowG49 = sheet.getRow(cellReferenceG49.getRow());
                        final HSSFCell cellG49 = rowG49.getCell(cellReferenceG49.getCol());
                        final CellReference cellReferenceJ49 = new CellReference("J49");
                        final HSSFRow rowJ49 = sheet.getRow(cellReferenceJ49.getRow());
                        final HSSFCell cellJ49 = rowJ49.getCell(cellReferenceJ49.getCol());
                        final CellReference cellReferenceM49 = new CellReference("M49");
                        final HSSFRow rowM49 = sheet.getRow(cellReferenceM49.getRow());
                        final HSSFCell cellM49 = rowM49.getCell(cellReferenceM49.getCol());
                        final double resultS49 = ((
                                cellG49.getNumericCellValue() + cellJ49.getNumericCellValue() + cellM49.getNumericCellValue()) / 3)
                                - cellD49.getNumericCellValue();
                        final CellReference cellReferenceD50 = new CellReference("D50");
                        final HSSFRow rowD50 = sheet.getRow(cellReferenceD50.getRow());
                        final HSSFCell cellD50 = rowD50.getCell(cellReferenceD50.getCol());
                        final CellReference cellReferenceG50 = new CellReference("G50");
                        final HSSFRow rowG50 = sheet.getRow(cellReferenceG50.getRow());
                        final HSSFCell cellG50 = rowG50.getCell(cellReferenceG50.getCol());
                        final CellReference cellReferenceJ50 = new CellReference("J50");
                        final HSSFRow rowJ50 = sheet.getRow(cellReferenceJ50.getRow());
                        final HSSFCell cellJ50 = rowJ50.getCell(cellReferenceJ50.getCol());
                        final CellReference cellReferenceM50 = new CellReference("M50");
                        final HSSFRow rowM50 = sheet.getRow(cellReferenceM50.getRow());
                        final HSSFCell cellM50 = rowM50.getCell(cellReferenceM50.getCol());
                        final double resultS50 = ((
                                cellG50.getNumericCellValue() + cellJ50.getNumericCellValue() + cellM50.getNumericCellValue()) / 3)
                                - cellD50.getNumericCellValue();
                        final CellReference cellReferenceD51 = new CellReference("D51");
                        final HSSFRow rowD51 = sheet.getRow(cellReferenceD51.getRow());
                        final HSSFCell cellD51 = rowD51.getCell(cellReferenceD51.getCol());
                        final CellReference cellReferenceG51 = new CellReference("G51");
                        final HSSFRow rowG51 = sheet.getRow(cellReferenceG51.getRow());
                        final HSSFCell cellG51 = rowG51.getCell(cellReferenceG51.getCol());
                        final CellReference cellReferenceJ51 = new CellReference("J51");
                        final HSSFRow rowJ51 = sheet.getRow(cellReferenceJ51.getRow());
                        final HSSFCell cellJ51 = rowJ51.getCell(cellReferenceJ51.getCol());
                        final CellReference cellReferenceM51 = new CellReference("M51");
                        final HSSFRow rowM51 = sheet.getRow(cellReferenceM51.getRow());
                        final HSSFCell cellM51 = rowM51.getCell(cellReferenceM51.getCol());
                        final double resultS51 = ((
                                cellG51.getNumericCellValue() + cellJ51.getNumericCellValue() + cellM51.getNumericCellValue()) / 3)
                                - cellD51.getNumericCellValue();
                        final CellReference cellReferenceD52 = new CellReference("D52");
                        final HSSFRow rowD52 = sheet.getRow(cellReferenceD52.getRow());
                        final HSSFCell cellD52 = rowD52.getCell(cellReferenceD52.getCol());
                        final CellReference cellReferenceG52 = new CellReference("G52");
                        final HSSFRow rowG52 = sheet.getRow(cellReferenceG52.getRow());
                        final HSSFCell cellG52 = rowG52.getCell(cellReferenceG52.getCol());
                        final CellReference cellReferenceJ52 = new CellReference("J52");
                        final HSSFRow rowJ52 = sheet.getRow(cellReferenceJ52.getRow());
                        final HSSFCell cellJ52 = rowJ52.getCell(cellReferenceJ52.getCol());
                        final CellReference cellReferenceM52 = new CellReference("M52");
                        final HSSFRow rowM52 = sheet.getRow(cellReferenceM52.getRow());
                        final HSSFCell cellM52 = rowM52.getCell(cellReferenceM52.getCol());
                        final double resultS52 = ((
                                cellG52.getNumericCellValue() + cellJ52.getNumericCellValue() + cellM52.getNumericCellValue()) / 3)
                                - cellD52.getNumericCellValue();
                        final CellReference cellReferenceD53 = new CellReference("D53");
                        final HSSFRow rowD53 = sheet.getRow(cellReferenceD53.getRow());
                        final HSSFCell cellD53 = rowD53.getCell(cellReferenceD53.getCol());
                        final CellReference cellReferenceG53 = new CellReference("G53");
                        final HSSFRow rowG53 = sheet.getRow(cellReferenceG53.getRow());
                        final HSSFCell cellG53 = rowG53.getCell(cellReferenceG53.getCol());
                        final CellReference cellReferenceJ53 = new CellReference("J53");
                        final HSSFRow rowJ53 = sheet.getRow(cellReferenceJ53.getRow());
                        final HSSFCell cellJ53 = rowJ53.getCell(cellReferenceJ53.getCol());
                        final CellReference cellReferenceM53 = new CellReference("M53");
                        final HSSFRow rowM53 = sheet.getRow(cellReferenceM53.getRow());
                        final HSSFCell cellM53 = rowM53.getCell(cellReferenceM53.getCol());
                        final double resultS53 = ((
                                cellG53.getNumericCellValue() + cellJ53.getNumericCellValue() + cellM53.getNumericCellValue()) / 3)
                                - cellD53.getNumericCellValue();
                        /*-----------------450-----------------*/
                        /*-----------------492-----------------*/
                        final CellReference cellReferenceD54 = new CellReference("D54");
                        final HSSFRow rowD54 = sheet.getRow(cellReferenceD54.getRow());
                        final HSSFCell cellD54 = rowD54.getCell(cellReferenceD54.getCol());
                        final CellReference cellReferenceG54 = new CellReference("G54");
                        final HSSFRow rowG54 = sheet.getRow(cellReferenceG54.getRow());
                        final HSSFCell cellG54 = rowG54.getCell(cellReferenceG54.getCol());
                        final CellReference cellReferenceJ54 = new CellReference("J54");
                        final HSSFRow rowJ54 = sheet.getRow(cellReferenceJ54.getRow());
                        final HSSFCell cellJ54 = rowJ54.getCell(cellReferenceJ54.getCol());
                        final CellReference cellReferenceM54 = new CellReference("M54");
                        final HSSFRow rowM54 = sheet.getRow(cellReferenceM54.getRow());
                        final HSSFCell cellM54 = rowM54.getCell(cellReferenceM54.getCol());
                        final double resultS54 = ((
                                cellG54.getNumericCellValue() + cellJ54.getNumericCellValue() + cellM54.getNumericCellValue()) / 3)
                                - cellD54.getNumericCellValue();
                        final CellReference cellReferenceD55 = new CellReference("D55");
                        final HSSFRow rowD55 = sheet.getRow(cellReferenceD55.getRow());
                        final HSSFCell cellD55 = rowD55.getCell(cellReferenceD55.getCol());
                        final CellReference cellReferenceG55 = new CellReference("G55");
                        final HSSFRow rowG55 = sheet.getRow(cellReferenceG55.getRow());
                        final HSSFCell cellG55 = rowG55.getCell(cellReferenceG55.getCol());
                        final CellReference cellReferenceJ55 = new CellReference("J55");
                        final HSSFRow rowJ55 = sheet.getRow(cellReferenceJ55.getRow());
                        final HSSFCell cellJ55 = rowJ55.getCell(cellReferenceJ55.getCol());
                        final CellReference cellReferenceM55 = new CellReference("M55");
                        final HSSFRow rowM55 = sheet.getRow(cellReferenceM55.getRow());
                        final HSSFCell cellM55 = rowM55.getCell(cellReferenceM55.getCol());
                        final double resultS55 = ((
                                cellG55.getNumericCellValue() + cellJ55.getNumericCellValue() + cellM55.getNumericCellValue()) / 3)
                                - cellD55.getNumericCellValue();
                        final CellReference cellReferenceD56 = new CellReference("D56");
                        final HSSFRow rowD56 = sheet.getRow(cellReferenceD56.getRow());
                        final HSSFCell cellD56 = rowD56.getCell(cellReferenceD56.getCol());
                        final CellReference cellReferenceG56 = new CellReference("G56");
                        final HSSFRow rowG56 = sheet.getRow(cellReferenceG56.getRow());
                        final HSSFCell cellG56 = rowG56.getCell(cellReferenceG56.getCol());
                        final CellReference cellReferenceJ56 = new CellReference("J56");
                        final HSSFRow rowJ56 = sheet.getRow(cellReferenceJ56.getRow());
                        final HSSFCell cellJ56 = rowJ56.getCell(cellReferenceJ56.getCol());
                        final CellReference cellReferenceM56 = new CellReference("M56");
                        final HSSFRow rowM56 = sheet.getRow(cellReferenceM56.getRow());
                        final HSSFCell cellM56 = rowM56.getCell(cellReferenceM56.getCol());
                        final double resultS56 = ((
                                cellG56.getNumericCellValue() + cellJ56.getNumericCellValue() + cellM56.getNumericCellValue()) / 3)
                                - cellD56.getNumericCellValue();
                        final CellReference cellReferenceD57 = new CellReference("D57");
                        final HSSFRow rowD57 = sheet.getRow(cellReferenceD57.getRow());
                        final HSSFCell cellD57 = rowD57.getCell(cellReferenceD57.getCol());
                        final CellReference cellReferenceG57 = new CellReference("G57");
                        final HSSFRow rowG57 = sheet.getRow(cellReferenceG57.getRow());
                        final HSSFCell cellG57 = rowG57.getCell(cellReferenceG57.getCol());
                        final CellReference cellReferenceJ57 = new CellReference("J57");
                        final HSSFRow rowJ57 = sheet.getRow(cellReferenceJ57.getRow());
                        final HSSFCell cellJ57 = rowJ57.getCell(cellReferenceJ57.getCol());
                        final CellReference cellReferenceM57 = new CellReference("M57");
                        final HSSFRow rowM57 = sheet.getRow(cellReferenceM57.getRow());
                        final HSSFCell cellM57 = rowM57.getCell(cellReferenceM57.getCol());
                        final double resultS57 = ((
                                cellG57.getNumericCellValue() + cellJ57.getNumericCellValue() + cellM57.getNumericCellValue()) / 3)
                                - cellD57.getNumericCellValue();
                        final CellReference cellReferenceD58 = new CellReference("D58");
                        final HSSFRow rowD58 = sheet.getRow(cellReferenceD58.getRow());
                        final HSSFCell cellD58 = rowD58.getCell(cellReferenceD58.getCol());
                        final CellReference cellReferenceG58 = new CellReference("G58");
                        final HSSFRow rowG58 = sheet.getRow(cellReferenceG58.getRow());
                        final HSSFCell cellG58 = rowG58.getCell(cellReferenceG58.getCol());
                        final CellReference cellReferenceJ58 = new CellReference("J58");
                        final HSSFRow rowJ58 = sheet.getRow(cellReferenceJ58.getRow());
                        final HSSFCell cellJ58 = rowJ58.getCell(cellReferenceJ58.getCol());
                        final CellReference cellReferenceM58 = new CellReference("M58");
                        final HSSFRow rowM58 = sheet.getRow(cellReferenceM58.getRow());
                        final HSSFCell cellM58 = rowM58.getCell(cellReferenceM58.getCol());
                        final double resultS58 = ((
                                cellG58.getNumericCellValue() + cellJ58.getNumericCellValue() + cellM58.getNumericCellValue()) / 3)
                                - cellD58.getNumericCellValue();
                        /*-----------------492-----------------*/
                        /*-----------------630-----------------*/
                        final CellReference cellReferenceD59 = new CellReference("D59");
                        final HSSFRow rowD59 = sheet.getRow(cellReferenceD59.getRow());
                        final HSSFCell cellD59 = rowD59.getCell(cellReferenceD59.getCol());
                        final CellReference cellReferenceG59 = new CellReference("G59");
                        final HSSFRow rowG59 = sheet.getRow(cellReferenceG59.getRow());
                        final HSSFCell cellG59 = rowG59.getCell(cellReferenceG59.getCol());
                        final CellReference cellReferenceJ59 = new CellReference("J59");
                        final HSSFRow rowJ59 = sheet.getRow(cellReferenceJ59.getRow());
                        final HSSFCell cellJ59 = rowJ59.getCell(cellReferenceJ59.getCol());
                        final CellReference cellReferenceM59 = new CellReference("M59");
                        final HSSFRow rowM59 = sheet.getRow(cellReferenceM59.getRow());
                        final HSSFCell cellM59 = rowM59.getCell(cellReferenceM59.getCol());
                        final double resultS59 = ((
                                cellG59.getNumericCellValue() + cellJ59.getNumericCellValue() + cellM59.getNumericCellValue()) / 3)
                                - cellD59.getNumericCellValue();
                        final CellReference cellReferenceD60 = new CellReference("D60");
                        final HSSFRow rowD60 = sheet.getRow(cellReferenceD60.getRow());
                        final HSSFCell cellD60 = rowD60.getCell(cellReferenceD60.getCol());
                        final CellReference cellReferenceG60 = new CellReference("G60");
                        final HSSFRow rowG60 = sheet.getRow(cellReferenceG60.getRow());
                        final HSSFCell cellG60 = rowG60.getCell(cellReferenceG60.getCol());
                        final CellReference cellReferenceJ60 = new CellReference("J60");
                        final HSSFRow rowJ60 = sheet.getRow(cellReferenceJ60.getRow());
                        final HSSFCell cellJ60 = rowJ60.getCell(cellReferenceJ60.getCol());
                        final CellReference cellReferenceM60 = new CellReference("M60");
                        final HSSFRow rowM60 = sheet.getRow(cellReferenceM60.getRow());
                        final HSSFCell cellM60 = rowM60.getCell(cellReferenceM60.getCol());
                        final double resultS60 = ((
                                cellG60.getNumericCellValue() + cellJ60.getNumericCellValue() + cellM60.getNumericCellValue()) / 3)
                                - cellD60.getNumericCellValue();
                        final CellReference cellReferenceD61 = new CellReference("D61");
                        final HSSFRow rowD61 = sheet.getRow(cellReferenceD61.getRow());
                        final HSSFCell cellD61 = rowD61.getCell(cellReferenceD61.getCol());
                        final CellReference cellReferenceG61 = new CellReference("G61");
                        final HSSFRow rowG61 = sheet.getRow(cellReferenceG61.getRow());
                        final HSSFCell cellG61 = rowG61.getCell(cellReferenceG61.getCol());
                        final CellReference cellReferenceJ61 = new CellReference("J61");
                        final HSSFRow rowJ61 = sheet.getRow(cellReferenceJ61.getRow());
                        final HSSFCell cellJ61 = rowJ61.getCell(cellReferenceJ61.getCol());
                        final CellReference cellReferenceM61 = new CellReference("M61");
                        final HSSFRow rowM61 = sheet.getRow(cellReferenceM61.getRow());
                        final HSSFCell cellM61 = rowM61.getCell(cellReferenceM61.getCol());
                        final double resultS61 = ((
                                cellG61.getNumericCellValue() + cellJ61.getNumericCellValue() + cellM61.getNumericCellValue()) / 3)
                                - cellD61.getNumericCellValue();
                        final CellReference cellReferenceD62 = new CellReference("D62");
                        final HSSFRow rowD62 = sheet.getRow(cellReferenceD62.getRow());
                        final HSSFCell cellD62 = rowD62.getCell(cellReferenceD62.getCol());
                        final CellReference cellReferenceG62 = new CellReference("G62");
                        final HSSFRow rowG62 = sheet.getRow(cellReferenceG62.getRow());
                        final HSSFCell cellG62 = rowG62.getCell(cellReferenceG62.getCol());
                        final CellReference cellReferenceJ62 = new CellReference("J62");
                        final HSSFRow rowJ62 = sheet.getRow(cellReferenceJ62.getRow());
                        final HSSFCell cellJ62 = rowJ62.getCell(cellReferenceJ62.getCol());
                        final CellReference cellReferenceM62 = new CellReference("M62");
                        final HSSFRow rowM62 = sheet.getRow(cellReferenceM62.getRow());
                        final HSSFCell cellM62 = rowM62.getCell(cellReferenceM62.getCol());
                        final double resultS62 = ((
                                cellG62.getNumericCellValue() + cellJ62.getNumericCellValue() + cellM62.getNumericCellValue()) / 3)
                                - cellD62.getNumericCellValue();
                        final CellReference cellReferenceD63 = new CellReference("D63");
                        final HSSFRow rowD63 = sheet.getRow(cellReferenceD63.getRow());
                        final HSSFCell cellD63 = rowD63.getCell(cellReferenceD63.getCol());
                        final CellReference cellReferenceG63 = new CellReference("G63");
                        final HSSFRow rowG63 = sheet.getRow(cellReferenceG63.getRow());
                        final HSSFCell cellG63 = rowG63.getCell(cellReferenceG63.getCol());
                        final CellReference cellReferenceJ63 = new CellReference("J63");
                        final HSSFRow rowJ63 = sheet.getRow(cellReferenceJ63.getRow());
                        final HSSFCell cellJ63 = rowJ63.getCell(cellReferenceJ63.getCol());
                        final CellReference cellReferenceM63 = new CellReference("M63");
                        final HSSFRow rowM63 = sheet.getRow(cellReferenceM63.getRow());
                        final HSSFCell cellM63 = rowM63.getCell(cellReferenceM63.getCol());
                        final double resultS63 = ((
                                cellG63.getNumericCellValue() + cellJ63.getNumericCellValue() + cellM63.getNumericCellValue()) / 3)
                                - cellD63.getNumericCellValue();
                        final ArrayList<Double> valueTolerance = new ArrayList<>();
                        valueTolerance.add(resultS44);
                        valueTolerance.add(resultS45);
                        valueTolerance.add(resultS46);
                        valueTolerance.add(resultS47);
                        valueTolerance.add(resultS48);
                        valueTolerance.add(resultS49);
                        valueTolerance.add(resultS50);
                        valueTolerance.add(resultS51);
                        valueTolerance.add(resultS52);
                        valueTolerance.add(resultS53);
                        valueTolerance.add(resultS54);
                        valueTolerance.add(resultS55);
                        valueTolerance.add(resultS56);
                        valueTolerance.add(resultS57);
                        valueTolerance.add(resultS58);
                        valueTolerance.add(resultS59);
                        valueTolerance.add(resultS60);
                        valueTolerance.add(resultS61);
                        valueTolerance.add(resultS62);
                        valueTolerance.add(resultS63);
                        final Double maxValueTolerance = Collections.max(valueTolerance);
                        final Double minValueTolerance = Collections.min(valueTolerance);
                        final double absMaxValueTolerance = Math.abs(maxValueTolerance);
                        final double absMinValueTolerance = Math.abs(minValueTolerance);
                        double biggerAbsNumber = 0;
                        if (absMaxValueTolerance > absMinValueTolerance) {
                            biggerAbsNumber = absMaxValueTolerance;
                        } else if (absMaxValueTolerance < absMinValueTolerance) {
                            biggerAbsNumber = absMinValueTolerance;
                        } else if (absMaxValueTolerance == absMinValueTolerance) {
                            biggerAbsNumber = absMaxValueTolerance;
                        }
                        String biggerAbsValueToleranceString = String.format("%.4f", biggerAbsNumber);
                        textFieldValueTolerance.setText(biggerAbsValueToleranceString);
                        if (biggerAbsNumber >= 0.015 || biggerAbsNumber <= -0.015) {
                            labelValueToleranceResult.setText("错误");
                            labelValueToleranceResult.setForeground(Color.red);
                        } else {
                            labelValueToleranceResult.setText("正确");
                            labelValueToleranceResult.setForeground(Color.green);
                        }
                        /*-----------------630-----------------*/
                        /*-----读取吸光度单元格-----*/
                        /*-----读取吸光度重复值单元格-----*/
                        final CellReference cellReferenceD70 = new CellReference("D70");
                        final HSSFRow rowD70 = sheet.getRow(cellReferenceD70.getRow());
                        final HSSFCell cellD70 = rowD70.getCell(cellReferenceD70.getCol());
                        final CellReference cellReferenceF70 = new CellReference("F70");
                        final HSSFRow rowF70 = sheet.getRow(cellReferenceF70.getRow());
                        final HSSFCell cellF70 = rowF70.getCell(cellReferenceF70.getCol());
                        final CellReference cellReferenceH70 = new CellReference("H70");
                        final HSSFRow rowH70 = sheet.getRow(cellReferenceH70.getRow());
                        final HSSFCell cellH70 = rowH70.getCell(cellReferenceH70.getCol());
                        final CellReference cellReferenceJ70 = new CellReference("J70");
                        final HSSFRow rowJ70 = sheet.getRow(cellReferenceJ70.getRow());
                        final HSSFCell cellJ70 = rowJ70.getCell(cellReferenceJ70.getCol());
                        final CellReference cellReferenceL70 = new CellReference("L70");
                        final HSSFRow rowL70 = sheet.getRow(cellReferenceL70.getRow());
                        final HSSFCell cellL70 = rowL70.getCell(cellReferenceL70.getCol());
                        final CellReference cellReferenceN70 = new CellReference("N70");
                        final HSSFRow rowN70 = sheet.getRow(cellReferenceN70.getRow());
                        final HSSFCell cellN70 = rowN70.getCell(cellReferenceN70.getCol());
                        final CellReference cellReferenceP70 = new CellReference("P70");
                        final HSSFRow rowP70 = sheet.getRow(cellReferenceP70.getRow());
                        final HSSFCell cellP70 = rowP70.getCell(cellReferenceP70.getCol());
                        final double averagep70 =
                                (
                                        cellD70.getNumericCellValue() +
                                                cellF70.getNumericCellValue() +
                                                cellH70.getNumericCellValue() +
                                                cellJ70.getNumericCellValue() +
                                                cellL70.getNumericCellValue() +
                                                cellN70.getNumericCellValue()
                                ) / 6;
                        final double dp70 = Math.pow((cellD70.getNumericCellValue() - averagep70), 2);
                        final double fp70 = Math.pow((cellF70.getNumericCellValue() - averagep70), 2);
                        final double hp70 = Math.pow((cellH70.getNumericCellValue() - averagep70), 2);
                        final double jp70 = Math.pow((cellJ70.getNumericCellValue() - averagep70), 2);
                        final double lp70 = Math.pow((cellL70.getNumericCellValue() - averagep70), 2);
                        final double np70 = Math.pow((cellN70.getNumericCellValue() - averagep70), 2);
                        double resultrsd = (Math.pow((dp70 + fp70 + hp70 + jp70 + lp70 + np70) / 5, 0.5)) / averagep70;
                        resultrsd = resultrsd * 100;
                        String resultrsdString = String.format("%.4f", resultrsd);
                        textFieldAbsorbanceRepeat.setText(resultrsdString);
                        if (resultrsd >= 0.1) {
                            labelAbsorbanceRepeatResult.setText("错误");
                            labelAbsorbanceRepeatResult.setForeground(Color.red);
                        } else {
                            labelAbsorbanceRepeatResult.setText("正确");
                            labelAbsorbanceRepeatResult.setForeground(Color.green);
                        }
                        /*-----读取吸光度重复值单元格-----*/
                        /*-----读取通道差异单元格-----*/
                        /*-----------------TunnalA-----------------*/
                        final CellReference cellReferenceD78 = new CellReference("D78");
                        final HSSFRow rowD78 = sheet.getRow(cellReferenceD78.getRow());
                        final HSSFCell cellD78 = rowD78.getCell(cellReferenceD78.getCol());
                        final CellReference cellReferenceG78 = new CellReference("G78");
                        final HSSFRow rowG78 = sheet.getRow(cellReferenceG78.getRow());
                        final HSSFCell cellG78 = rowG78.getCell(cellReferenceG78.getCol());
                        final CellReference cellReferenceJ78 = new CellReference("J78");
                        final HSSFRow rowJ78 = sheet.getRow(cellReferenceJ78.getRow());
                        final HSSFCell cellJ78 = rowJ78.getCell(cellReferenceJ78.getCol());
                        final CellReference cellReferenceM78 = new CellReference("M78");
                        final HSSFRow rowM78 = sheet.getRow(cellReferenceM78.getRow());
                        final HSSFCell cellM78 = rowM78.getCell(cellReferenceM78.getCol());
                        final CellReference cellReferenceP78 = new CellReference("P78");
                        final HSSFRow rowP78 = sheet.getRow(cellReferenceP78.getRow());
                        final HSSFCell cellP78 = rowP78.getCell(cellReferenceP78.getCol());
                        final CellReference cellReferenceS78 = new CellReference("S78");
                        final HSSFRow rowS78 = sheet.getRow(cellReferenceS78.getRow());
                        final HSSFCell cellS78 = rowS78.getCell(cellReferenceS78.getCol());
                        final CellReference cellReferenceV78 = new CellReference("V78");
                        final HSSFRow rowV78 = sheet.getRow(cellReferenceV78.getRow());
                        final HSSFCell cellV78 = rowV78.getCell(cellReferenceV78.getCol());
                        final CellReference cellReferenceZ78 = new CellReference("Z78");
                        final HSSFRow rowZ78 = sheet.getRow(cellReferenceZ78.getRow());
                        final HSSFCell cellZ78 = rowZ78.getCell(cellReferenceZ78.getCol());
                        final ArrayList<Double> tunnalA = new ArrayList<>();
                        tunnalA.add(cellD78.getNumericCellValue());
                        tunnalA.add(cellG78.getNumericCellValue());
                        tunnalA.add(cellJ78.getNumericCellValue());
                        tunnalA.add(cellM78.getNumericCellValue());
                        tunnalA.add(cellP78.getNumericCellValue());
                        tunnalA.add(cellS78.getNumericCellValue());
                        double maxTunnalA = Collections.max(tunnalA);
                        double minTunnalA = Collections.min(tunnalA);
                        maxTunnalA = cellV78.getNumericCellValue();
                        minTunnalA = cellZ78.getNumericCellValue();
                        /*-----------------TunnalA-----------------*/
                        /*-----------------TunnalB-----------------*/
                        final CellReference cellReferenceD79 = new CellReference("D79");
                        final HSSFRow rowD79 = sheet.getRow(cellReferenceD79.getRow());
                        final HSSFCell cellD79 = rowD79.getCell(cellReferenceD79.getCol());
                        final CellReference cellReferenceG79 = new CellReference("G79");
                        final HSSFRow rowG79 = sheet.getRow(cellReferenceG79.getRow());
                        final HSSFCell cellG79 = rowG79.getCell(cellReferenceG79.getCol());
                        final CellReference cellReferenceJ79 = new CellReference("J79");
                        final HSSFRow rowJ79 = sheet.getRow(cellReferenceJ79.getRow());
                        final HSSFCell cellJ79 = rowJ79.getCell(cellReferenceJ79.getCol());
                        final CellReference cellReferenceM79 = new CellReference("M79");
                        final HSSFRow rowM79 = sheet.getRow(cellReferenceM79.getRow());
                        final HSSFCell cellM79 = rowM79.getCell(cellReferenceM79.getCol());
                        final CellReference cellReferenceP79 = new CellReference("P79");
                        final HSSFRow rowP79 = sheet.getRow(cellReferenceP79.getRow());
                        final HSSFCell cellP79 = rowP79.getCell(cellReferenceP79.getCol());
                        final CellReference cellReferenceS79 = new CellReference("S79");
                        final HSSFRow rowS79 = sheet.getRow(cellReferenceS79.getRow());
                        final HSSFCell cellS79 = rowS79.getCell(cellReferenceS79.getCol());
                        final CellReference cellReferenceV79 = new CellReference("V79");
                        final HSSFRow rowV79 = sheet.getRow(cellReferenceV79.getRow());
                        final HSSFCell cellV79 = rowV79.getCell(cellReferenceV79.getCol());
                        final CellReference cellReferenceZ79 = new CellReference("Z79");
                        final HSSFRow rowZ79 = sheet.getRow(cellReferenceZ79.getRow());
                        final HSSFCell cellZ79 = rowZ79.getCell(cellReferenceZ79.getCol());
                        final ArrayList<Double> tunnalB = new ArrayList<>();
                        tunnalB.add(cellD79.getNumericCellValue());
                        tunnalB.add(cellG79.getNumericCellValue());
                        tunnalB.add(cellJ79.getNumericCellValue());
                        tunnalB.add(cellM79.getNumericCellValue());
                        tunnalB.add(cellP79.getNumericCellValue());
                        tunnalB.add(cellS79.getNumericCellValue());
                        double maxtunnalB = Collections.max(tunnalB);
                        double mintunnalB = Collections.min(tunnalB);
                        maxtunnalB = cellV79.getNumericCellValue();
                        mintunnalB = cellZ79.getNumericCellValue();
                        /*-----------------TunnalB-----------------*/
                        /*-----------------Amax-----------------*/
                        final ArrayList<Double> Amax = new ArrayList<>();
                        Amax.add(maxTunnalA);
                        Amax.add(maxtunnalB);
                        double resultAmax = Collections.max(Amax);
                        /*-----------------Amax-----------------*/
                        /*-----------------Amin-----------------*/
                        final ArrayList<Double> Amin = new ArrayList<>();
                        Amin.add(minTunnalA);
                        Amin.add(mintunnalB);
                        double resultAmin = Collections.min(Amin);
                        /*-----------------Amin-----------------*/
                        String resultAmaxAminString = String.format("%.4f", resultAmax - resultAmin);
                        textFieldChannelDifference.setText(resultAmaxAminString);
                        if ((resultAmax - resultAmin) > 0.02) {
                            labelChannelDifferenceResult.setText("错误");
                            labelChannelDifferenceResult.setForeground(Color.red);
                        } else {
                            labelChannelDifferenceResult.setText("正确");
                            labelChannelDifferenceResult.setForeground(Color.green);
                        }
                        /*-----读取通道差异单元格-----*/
                        /*----------读取U盘XLS文件的数据----------*/
                    }
                    /*-----中文报告-----*/
                    /*-----中文报告-----*/
                }
            } else {
                textFieldType.setText("");
                textFieldNumber.setText("");
                textFieldValueStability.setText("");
                textFieldValueTolerance.setText("");
                textFieldAbsorbanceRepeat.setText("");
                textFieldChannelDifference.setText("");
                labelTunnalMinResult.setText("");
                labelTunnalMaxResult.setText("");
                labelValueStabilityResult.setText("");
                labelValueToleranceResult.setText("");
                labelAbsorbanceRepeatResult.setText("");
                labelChannelDifferenceResult.setText("");
            }
            /*-----方法2-----*/
        }
        /*-----获取所有文件名-----*/

        //查找资源--生产者使用
        public synchronized void searchFile() {
            if (flag) {
                try {
                    wait();
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
            dirs = File.listRoots();
            if (dirs.length > count) {
                flag = true;
                notify();
            }
        }

        //消费资源--消费者使用
        public synchronized void readFile() {
            if (!flag) {
                try {
                    wait();
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
            if (dirs.length > count) {
                for (i = count; i < dirs.length; i++) {
                    try {
                        getAllFiles(dirs[i]);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                flag = false;
                notify();
            }
        }
    }

    //消费者
    public static class ConsumerUSBRoot implements Runnable {
        public boolean isClose = false;
        private ResFile rf = null;

        public ConsumerUSBRoot(ResFile rf) {
            this.rf = rf;
        }

        @Override
        public void run() {
            while (true) {
                if (isClose) {
                    //System.out.println("关闭USB消费者监听...");
                    break;
                }
                rf.readFile();
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    //生产者
    public static class ProducerUSBRoot implements Runnable {
        public boolean isClose = false;
        private ResFile rf = null;

        public ProducerUSBRoot(ResFile rf) {
            this.rf = rf;
        }

        @Override
        public void run() {
            while (true) {
                rf.searchFile();
                if (isClose) {
                    //System.out.println("关闭USB生产者监听...");
                    break;
                }
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
