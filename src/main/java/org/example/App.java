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
 * 2.消费者：在没有判断出有插入设备时，处于等待状态；若有，则查找设备中是否包含指定文件，有则关机。
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
        final JTextField textFieldNumber = new JTextField(50); //设置机身号文本框
        final JTextField textFieldTunnalMin = new JTextField(50); //设置通道最小值提示文本框
        final JTextField textFieldTunnalMax = new JTextField(50); //设置通道最大值提示文本框
        final JTextField textFieldValueStability = new JTextField(50); //设置示值稳定性提示文本框
        final JTextField textFieldValueTolerance = new JTextField(50); //设置吸光度示值误差提示文本框
        final JTextField textFieldAbsorbanceRepeat = new JTextField(50); //设置吸光度重复值提示文本框
        final JTextField textFieldSensitive = new JTextField(50); //设置灵敏度提示文本框
        final JTextField textFieldChannelDifference = new JTextField(50); //设置通道差异提示文本框
        final JTextField textFieldSoftwareFolder = new JTextField(50); //设置软件提示文本框
        final JTextField textFieldSystemGHO = new JTextField(50); //设置系统文件提示文本框
        Font fontLabel = new Font("宋体", Font.PLAIN, 15);//设置标签字体
        Font fontNumber = new Font("宋体", Font.BOLD, 60);//设置型号和机身号字体
        Font fontResult = new Font("宋体", Font.PLAIN, 14);//设置结果字体格式
        final JLabel labelTunnalMinResult = new JLabel(); //设置通道最小值结果显示标签字样
        final JLabel labelTunnalMaxResult = new JLabel(); //设置通道最大值结果显示标签字样
        final JLabel labelValueStabilityResult = new JLabel(); //设置示值稳定性结果显示标签字样
        final JLabel labelValueToleranceResult = new JLabel(); //设置吸光度示值误差结果显示标签字样
        final JLabel labelAbsorbanceRepeatResult = new JLabel(); //设置吸光度重复值结果显示标签字样
        final JLabel labelSensitiveResult = new JLabel(); //设置灵敏度结果显示标签字样
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
            /*---------------机身号文本框---------------*/
            textFieldNumber.setBounds(25, 71, 435, 55); //设置位置大小
            textFieldNumber.setEditable(false);//设置不可编辑
            textFieldNumber.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldNumber); //添加标签
            /*---------------机身号文本框---------------*/
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
            /*---------------灵敏度标签---------------*/
            final JLabel labelSensitive = new JLabel("灵敏度（>0.05）："); //设置目录字样
            labelSensitive.setBounds(25, 241, 255, 25); //设置位置大小
            labelSensitive.setFont(fontLabel);
            this.add(labelSensitive); //添加标签
            /*---------------灵敏度标签---------------*/
            /*---------------灵敏度文本框---------------*/
            textFieldSensitive.setBounds(235, 237, 150, 30); //设置位置大小
            textFieldSensitive.setEditable(false);//设置不可编辑
            textFieldSensitive.setBackground(Color.white);//设置文本框背景色为白色
            this.add(textFieldSensitive); //添加标签
            /*---------------灵敏度文本框---------------*/
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
//            labelValueStabilityResult.setText("正确");//测试显示效果
            this.add(labelValueStabilityResult);//添加标签
            /*---------------设置示值稳定性结果显示标签字样---------------*/
            /*---------------设置吸光度示值误差结果显示标签字样---------------*/
            labelValueToleranceResult.setBounds(395, 170, 30, 25); //设置位置大小
            labelValueToleranceResult.setFont(fontLabel);//设置字体
//            labelValueToleranceResult.setText("正确");//测试显示效果
            this.add(labelValueToleranceResult);//添加标签
            /*---------------设置吸光度示值误差结果显示标签字样---------------*/
            /*---------------设置吸光度重复值结果显示标签字样---------------*/
            labelAbsorbanceRepeatResult.setBounds(395, 206, 30, 25); //设置位置大小
            labelAbsorbanceRepeatResult.setFont(fontLabel);//设置字体
//            labelAbsorbanceRepeatResult.setText("正确");//测试显示效果
            this.add(labelAbsorbanceRepeatResult);//添加标签
            /*---------------设置吸光度重复值结果显示标签字样---------------*/
            /*---------------设置灵敏度结果显示标签字样---------------*/
            labelSensitiveResult.setBounds(395, 241, 30, 25); //设置位置大小
            labelSensitiveResult.setFont(fontLabel);//设置字体
//            labelSensitiveResult.setText("正确");//测试显示效果
            this.add(labelSensitiveResult);//添加标签
            /*---------------设置灵敏度结果显示标签字样---------------*/
            /*---------------设置通道差异结果显示标签字样---------------*/
            labelChannelDifferenceResult.setBounds(395, 276, 30, 25); //设置位置大小
            labelChannelDifferenceResult.setFont(fontLabel);//设置字体
//            labelChannelDifferenceResult.setText("正确");//测试显示效果
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
//            labelSystemGHOResult.setText("测试");//测试显示效果
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
            /*-----实验方法-----*/
            if (dirAllStrArr.size() <= 2) {
                textFieldSoftwareFolder.setText("");
                labelSoftwareFolderResult.setIcon(null);
                textFieldSystemGHO.setText("");
                textFieldType.setText("");
                textFieldNumber.setText("");
                textFieldValueStability.setText("");
                textFieldValueTolerance.setText("");
                textFieldAbsorbanceRepeat.setText("");
                textFieldSensitive.setText("");
                textFieldChannelDifference.setText("");
                labelTunnalMinResult.setText("");
                labelTunnalMaxResult.setText("");
                labelValueStabilityResult.setText("");
                labelValueToleranceResult.setText("");
                labelAbsorbanceRepeatResult.setText("");
                labelSensitiveResult.setText("");
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

//                            System.out.println("file: " + f); //打印完整路径
//                            System.out.println(f.getName()+"\n===================="); //打印文件名
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
                    /*----------英文----------*/
                    Pattern pattern = Pattern.compile("[A-Z](.*?)\\d$");
                    /*-----读取型号单元格数据-----*/
                    final CellReference cellReferenceG3 = new CellReference("G3");
                    final HSSFRow rowG3 = sheet.getRow(cellReferenceG3.getRow());
                    final HSSFCell cellG3 = rowG3.getCell(cellReferenceG3.getCol());
                    String enProductType = cellG3.getStringCellValue();
                    final Matcher enmatcherProductType = pattern.matcher(enProductType.replaceAll("\\s", ""));
                    /*-----读取型号单元格数据-----*/
                    /*-----读取机身号单元格数据-----*/
                    final CellReference cellReferenceG4 = new CellReference("G4");
                    final HSSFRow rowG4 = sheet.getRow(cellReferenceG4.getRow());
                    final HSSFCell cellG4 = rowG4.getCell(cellReferenceG4.getCol());
                    String enProductNumber = cellG4.getStringCellValue();
                    Matcher enmatcherProductNumber = pattern.matcher(enProductNumber.replaceAll("\\s", ""));
                    /*-----读取机身号单元格数据-----*/
                    /*----------英文----------*/
//                  System.out.println("机身号为：" + ProductNumber);
//                  System.out.println("============================");
                    textFieldType.setFont(fontNumber);
                    textFieldType.setText(enmatcherProductType.group());
                    textFieldNumber.setFont(fontNumber);
                    textFieldNumber.setText(enmatcherProductNumber.group());
                    /*----------读取U盘XLS文件的数据----------*/
                    /*-----读取示值稳定性单元格-----*/
                    final CellReference cellReferenceA22 = new CellReference("A22");
                    final HSSFRow rowA22 = sheet.getRow(cellReferenceA22.getRow());
                    final HSSFCell cellA22 = rowA22.getCell(cellReferenceA22.getCol());
                    final CellReference cellReferenceF22 = new CellReference("F22");
                    final HSSFRow rowF22 = sheet.getRow(cellReferenceF22.getRow());
                    final HSSFCell cellF22 = rowF22.getCell(cellReferenceF22.getCol());
                    final CellReference cellReferenceK22 = new CellReference("K22");
                    final HSSFRow rowK22 = sheet.getRow(cellReferenceK22.getRow());
                    final HSSFCell cellK22 = rowK22.getCell(cellReferenceK22.getCol());
                    final double resultP22 = (Math.max(cellF22.getNumericCellValue(), cellK22.getNumericCellValue())) - cellA22.getNumericCellValue();
                    String resultP22String = String.format("%.4f", resultP22);
                    textFieldValueStability.setFont(fontResult);//设置字体
                    textFieldValueStability.setText(resultP22String);
                    if (resultP22 >= 0.002 || resultP22 <= -0.002) {
                        labelValueStabilityResult.setText("错误");
                        labelValueStabilityResult.setForeground(Color.red);
                    } else {
                        labelValueStabilityResult.setText("正确");
                        labelValueStabilityResult.setForeground(Color.green);
                    }
                    /*-----读取示值稳定性单元格-----*/
                    /*-----读取吸光度单元格-----*/
                    /*-----------------405-----------------*/
                    final CellReference cellReferenceD27 = new CellReference("D27");
                    final HSSFRow rowD27 = sheet.getRow(cellReferenceD27.getRow());
                    final HSSFCell cellD27 = rowD27.getCell(cellReferenceD27.getCol());
                    final CellReference cellReferenceG27 = new CellReference("G27");
                    final HSSFRow rowG27 = sheet.getRow(cellReferenceG27.getRow());
                    final HSSFCell cellG27 = rowG27.getCell(cellReferenceG27.getCol());
                    final CellReference cellReferenceJ27 = new CellReference("J27");
                    final HSSFRow rowJ27 = sheet.getRow(cellReferenceJ27.getRow());
                    final HSSFCell cellJ27 = rowJ27.getCell(cellReferenceJ27.getCol());
                    final CellReference cellReferenceM27 = new CellReference("M27");
                    final HSSFRow rowM27 = sheet.getRow(cellReferenceM27.getRow());
                    final HSSFCell cellM27 = rowM27.getCell(cellReferenceM27.getCol());
                    final double resultS27 = ((
                            cellG27.getNumericCellValue() + cellJ27.getNumericCellValue() + cellM27.getNumericCellValue()) / 3)
                            - cellD27.getNumericCellValue();
                    final CellReference cellReferenceD28 = new CellReference("D28");
                    final HSSFRow rowD28 = sheet.getRow(cellReferenceD28.getRow());
                    final HSSFCell cellD28 = rowD28.getCell(cellReferenceD28.getCol());
                    final CellReference cellReferenceG28 = new CellReference("G28");
                    final HSSFRow rowG28 = sheet.getRow(cellReferenceG28.getRow());
                    final HSSFCell cellG28 = rowG28.getCell(cellReferenceG28.getCol());
                    final CellReference cellReferenceJ28 = new CellReference("J28");
                    final HSSFRow rowJ28 = sheet.getRow(cellReferenceJ28.getRow());
                    final HSSFCell cellJ28 = rowJ28.getCell(cellReferenceJ28.getCol());
                    final CellReference cellReferenceM28 = new CellReference("M28");
                    final HSSFRow rowM28 = sheet.getRow(cellReferenceM28.getRow());
                    final HSSFCell cellM28 = rowM28.getCell(cellReferenceM28.getCol());
                    final double resultS28 = ((
                            cellG28.getNumericCellValue() + cellJ28.getNumericCellValue() + cellM28.getNumericCellValue()) / 3)
                            - cellD28.getNumericCellValue();
                    final CellReference cellReferenceD29 = new CellReference("D29");
                    final HSSFRow rowD29 = sheet.getRow(cellReferenceD29.getRow());
                    final HSSFCell cellD29 = rowD29.getCell(cellReferenceD29.getCol());
                    final CellReference cellReferenceG29 = new CellReference("G29");
                    final HSSFRow rowG29 = sheet.getRow(cellReferenceG29.getRow());
                    final HSSFCell cellG29 = rowG29.getCell(cellReferenceG29.getCol());
                    final CellReference cellReferenceJ29 = new CellReference("J29");
                    final HSSFRow rowJ29 = sheet.getRow(cellReferenceJ29.getRow());
                    final HSSFCell cellJ29 = rowJ29.getCell(cellReferenceJ29.getCol());
                    final CellReference cellReferenceM29 = new CellReference("M29");
                    final HSSFRow rowM29 = sheet.getRow(cellReferenceM29.getRow());
                    final HSSFCell cellM29 = rowM29.getCell(cellReferenceM29.getCol());
                    final double resultS29 = ((
                            cellG29.getNumericCellValue() + cellJ29.getNumericCellValue() + cellM29.getNumericCellValue()) / 3)
                            - cellD29.getNumericCellValue();
                    final CellReference cellReferenceD30 = new CellReference("D30");
                    final HSSFRow rowD30 = sheet.getRow(cellReferenceD30.getRow());
                    final HSSFCell cellD30 = rowD30.getCell(cellReferenceD30.getCol());
                    final CellReference cellReferenceG30 = new CellReference("G30");
                    final HSSFRow rowG30 = sheet.getRow(cellReferenceG30.getRow());
                    final HSSFCell cellG30 = rowG30.getCell(cellReferenceG30.getCol());
                    final CellReference cellReferenceJ30 = new CellReference("J30");
                    final HSSFRow rowJ30 = sheet.getRow(cellReferenceJ30.getRow());
                    final HSSFCell cellJ30 = rowJ30.getCell(cellReferenceJ30.getCol());
                    final CellReference cellReferenceM30 = new CellReference("M30");
                    final HSSFRow rowM30 = sheet.getRow(cellReferenceM30.getRow());
                    final HSSFCell cellM30 = rowM30.getCell(cellReferenceM30.getCol());
                    final double resultS30 = ((
                            cellG30.getNumericCellValue() + cellJ30.getNumericCellValue() + cellM30.getNumericCellValue()) / 3)
                            - cellD30.getNumericCellValue();
                    final CellReference cellReferenceD31 = new CellReference("D31");
                    final HSSFRow rowD31 = sheet.getRow(cellReferenceD31.getRow());
                    final HSSFCell cellD31 = rowD31.getCell(cellReferenceD31.getCol());
                    final CellReference cellReferenceG31 = new CellReference("G31");
                    final HSSFRow rowG31 = sheet.getRow(cellReferenceG31.getRow());
                    final HSSFCell cellG31 = rowG31.getCell(cellReferenceG31.getCol());
                    final CellReference cellReferenceJ31 = new CellReference("J31");
                    final HSSFRow rowJ31 = sheet.getRow(cellReferenceJ31.getRow());
                    final HSSFCell cellJ31 = rowJ31.getCell(cellReferenceJ31.getCol());
                    final CellReference cellReferenceM31 = new CellReference("M31");
                    final HSSFRow rowM31 = sheet.getRow(cellReferenceM31.getRow());
                    final HSSFCell cellM31 = rowM31.getCell(cellReferenceM31.getCol());
                    final double resultS31 = ((
                            cellG31.getNumericCellValue() + cellJ31.getNumericCellValue() + cellM31.getNumericCellValue()) / 3)
                            - cellD31.getNumericCellValue();
                    /*-----------------405-----------------*/
                    /*-----------------450-----------------*/
                    final CellReference cellReferenceD32 = new CellReference("D32");
                    final HSSFRow rowD32 = sheet.getRow(cellReferenceD32.getRow());
                    final HSSFCell cellD32 = rowD32.getCell(cellReferenceD32.getCol());
                    final CellReference cellReferenceG32 = new CellReference("G32");
                    final HSSFRow rowG32 = sheet.getRow(cellReferenceG32.getRow());
                    final HSSFCell cellG32 = rowG32.getCell(cellReferenceG32.getCol());
                    final CellReference cellReferenceJ32 = new CellReference("J32");
                    final HSSFRow rowJ32 = sheet.getRow(cellReferenceJ32.getRow());
                    final HSSFCell cellJ32 = rowJ32.getCell(cellReferenceJ32.getCol());
                    final CellReference cellReferenceM32 = new CellReference("M32");
                    final HSSFRow rowM32 = sheet.getRow(cellReferenceM32.getRow());
                    final HSSFCell cellM32 = rowM32.getCell(cellReferenceM32.getCol());
                    final double resultS32 = ((
                            cellG32.getNumericCellValue() + cellJ32.getNumericCellValue() + cellM32.getNumericCellValue()) / 3)
                            - cellD32.getNumericCellValue();
                    final CellReference cellReferenceD33 = new CellReference("D33");
                    final HSSFRow rowD33 = sheet.getRow(cellReferenceD33.getRow());
                    final HSSFCell cellD33 = rowD33.getCell(cellReferenceD33.getCol());
                    final CellReference cellReferenceG33 = new CellReference("G33");
                    final HSSFRow rowG33 = sheet.getRow(cellReferenceG33.getRow());
                    final HSSFCell cellG33 = rowG33.getCell(cellReferenceG33.getCol());
                    final CellReference cellReferenceJ33 = new CellReference("J33");
                    final HSSFRow rowJ33 = sheet.getRow(cellReferenceJ33.getRow());
                    final HSSFCell cellJ33 = rowJ33.getCell(cellReferenceJ33.getCol());
                    final CellReference cellReferenceM33 = new CellReference("M33");
                    final HSSFRow rowM33 = sheet.getRow(cellReferenceM33.getRow());
                    final HSSFCell cellM33 = rowM33.getCell(cellReferenceM33.getCol());
                    final double resultS33 = ((
                            cellG33.getNumericCellValue() + cellJ33.getNumericCellValue() + cellM33.getNumericCellValue()) / 3)
                            - cellD33.getNumericCellValue();
                    final CellReference cellReferenceD34 = new CellReference("D34");
                    final HSSFRow rowD34 = sheet.getRow(cellReferenceD34.getRow());
                    final HSSFCell cellD34 = rowD34.getCell(cellReferenceD34.getCol());
                    final CellReference cellReferenceG34 = new CellReference("G34");
                    final HSSFRow rowG34 = sheet.getRow(cellReferenceG34.getRow());
                    final HSSFCell cellG34 = rowG34.getCell(cellReferenceG34.getCol());
                    final CellReference cellReferenceJ34 = new CellReference("J34");
                    final HSSFRow rowJ34 = sheet.getRow(cellReferenceJ34.getRow());
                    final HSSFCell cellJ34 = rowJ34.getCell(cellReferenceJ34.getCol());
                    final CellReference cellReferenceM34 = new CellReference("M34");
                    final HSSFRow rowM34 = sheet.getRow(cellReferenceM34.getRow());
                    final HSSFCell cellM34 = rowM34.getCell(cellReferenceM34.getCol());
                    final double resultS34 = ((
                            cellG34.getNumericCellValue() + cellJ34.getNumericCellValue() + cellM34.getNumericCellValue()) / 3)
                            - cellD34.getNumericCellValue();
                    final CellReference cellReferenceD35 = new CellReference("D35");
                    final HSSFRow rowD35 = sheet.getRow(cellReferenceD35.getRow());
                    final HSSFCell cellD35 = rowD35.getCell(cellReferenceD35.getCol());
                    final CellReference cellReferenceG35 = new CellReference("G35");
                    final HSSFRow rowG35 = sheet.getRow(cellReferenceG35.getRow());
                    final HSSFCell cellG35 = rowG35.getCell(cellReferenceG35.getCol());
                    final CellReference cellReferenceJ35 = new CellReference("J35");
                    final HSSFRow rowJ35 = sheet.getRow(cellReferenceJ35.getRow());
                    final HSSFCell cellJ35 = rowJ35.getCell(cellReferenceJ35.getCol());
                    final CellReference cellReferenceM35 = new CellReference("M35");
                    final HSSFRow rowM35 = sheet.getRow(cellReferenceM35.getRow());
                    final HSSFCell cellM35 = rowM35.getCell(cellReferenceM35.getCol());
                    final double resultS35 = ((
                            cellG35.getNumericCellValue() + cellJ35.getNumericCellValue() + cellM35.getNumericCellValue()) / 3)
                            - cellD35.getNumericCellValue();
                    final CellReference cellReferenceD36 = new CellReference("D36");
                    final HSSFRow rowD36 = sheet.getRow(cellReferenceD36.getRow());
                    final HSSFCell cellD36 = rowD36.getCell(cellReferenceD36.getCol());
                    final CellReference cellReferenceG36 = new CellReference("G36");
                    final HSSFRow rowG36 = sheet.getRow(cellReferenceG36.getRow());
                    final HSSFCell cellG36 = rowG36.getCell(cellReferenceG36.getCol());
                    final CellReference cellReferenceJ36 = new CellReference("J36");
                    final HSSFRow rowJ36 = sheet.getRow(cellReferenceJ36.getRow());
                    final HSSFCell cellJ36 = rowJ36.getCell(cellReferenceJ36.getCol());
                    final CellReference cellReferenceM36 = new CellReference("M36");
                    final HSSFRow rowM36 = sheet.getRow(cellReferenceM36.getRow());
                    final HSSFCell cellM36 = rowM36.getCell(cellReferenceM36.getCol());
                    final double resultS36 = ((
                            cellG36.getNumericCellValue() + cellJ36.getNumericCellValue() + cellM36.getNumericCellValue()) / 3)
                            - cellD36.getNumericCellValue();
                    /*-----------------450-----------------*/
                    /*-----------------492-----------------*/
                    final CellReference cellReferenceD37 = new CellReference("D37");
                    final HSSFRow rowD37 = sheet.getRow(cellReferenceD37.getRow());
                    final HSSFCell cellD37 = rowD37.getCell(cellReferenceD37.getCol());
                    final CellReference cellReferenceG37 = new CellReference("G37");
                    final HSSFRow rowG37 = sheet.getRow(cellReferenceG37.getRow());
                    final HSSFCell cellG37 = rowG37.getCell(cellReferenceG37.getCol());
                    final CellReference cellReferenceJ37 = new CellReference("J37");
                    final HSSFRow rowJ37 = sheet.getRow(cellReferenceJ37.getRow());
                    final HSSFCell cellJ37 = rowJ37.getCell(cellReferenceJ37.getCol());
                    final CellReference cellReferenceM37 = new CellReference("M37");
                    final HSSFRow rowM37 = sheet.getRow(cellReferenceM37.getRow());
                    final HSSFCell cellM37 = rowM37.getCell(cellReferenceM37.getCol());
                    final double resultS37 = ((
                            cellG37.getNumericCellValue() + cellJ37.getNumericCellValue() + cellM37.getNumericCellValue()) / 3)
                            - cellD37.getNumericCellValue();
                    final CellReference cellReferenceD38 = new CellReference("D38");
                    final HSSFRow rowD38 = sheet.getRow(cellReferenceD38.getRow());
                    final HSSFCell cellD38 = rowD38.getCell(cellReferenceD38.getCol());
                    final CellReference cellReferenceG38 = new CellReference("G38");
                    final HSSFRow rowG38 = sheet.getRow(cellReferenceG38.getRow());
                    final HSSFCell cellG38 = rowG38.getCell(cellReferenceG38.getCol());
                    final CellReference cellReferenceJ38 = new CellReference("J38");
                    final HSSFRow rowJ38 = sheet.getRow(cellReferenceJ38.getRow());
                    final HSSFCell cellJ38 = rowJ38.getCell(cellReferenceJ38.getCol());
                    final CellReference cellReferenceM38 = new CellReference("M38");
                    final HSSFRow rowM38 = sheet.getRow(cellReferenceM38.getRow());
                    final HSSFCell cellM38 = rowM38.getCell(cellReferenceM38.getCol());
                    final double resultS38 = ((
                            cellG38.getNumericCellValue() + cellJ38.getNumericCellValue() + cellM38.getNumericCellValue()) / 3)
                            - cellD38.getNumericCellValue();
                    final CellReference cellReferenceD39 = new CellReference("D39");
                    final HSSFRow rowD39 = sheet.getRow(cellReferenceD39.getRow());
                    final HSSFCell cellD39 = rowD39.getCell(cellReferenceD39.getCol());
                    final CellReference cellReferenceG39 = new CellReference("G39");
                    final HSSFRow rowG39 = sheet.getRow(cellReferenceG39.getRow());
                    final HSSFCell cellG39 = rowG39.getCell(cellReferenceG39.getCol());
                    final CellReference cellReferenceJ39 = new CellReference("J39");
                    final HSSFRow rowJ39 = sheet.getRow(cellReferenceJ39.getRow());
                    final HSSFCell cellJ39 = rowJ39.getCell(cellReferenceJ39.getCol());
                    final CellReference cellReferenceM39 = new CellReference("M39");
                    final HSSFRow rowM39 = sheet.getRow(cellReferenceM39.getRow());
                    final HSSFCell cellM39 = rowM39.getCell(cellReferenceM39.getCol());
                    final double resultS39 = ((
                            cellG39.getNumericCellValue() + cellJ39.getNumericCellValue() + cellM39.getNumericCellValue()) / 3)
                            - cellD39.getNumericCellValue();
                    final CellReference cellReferenceD40 = new CellReference("D40");
                    final HSSFRow rowD40 = sheet.getRow(cellReferenceD40.getRow());
                    final HSSFCell cellD40 = rowD40.getCell(cellReferenceD40.getCol());
                    final CellReference cellReferenceG40 = new CellReference("G40");
                    final HSSFRow rowG40 = sheet.getRow(cellReferenceG40.getRow());
                    final HSSFCell cellG40 = rowG40.getCell(cellReferenceG40.getCol());
                    final CellReference cellReferenceJ40 = new CellReference("J40");
                    final HSSFRow rowJ40 = sheet.getRow(cellReferenceJ40.getRow());
                    final HSSFCell cellJ40 = rowJ40.getCell(cellReferenceJ40.getCol());
                    final CellReference cellReferenceM40 = new CellReference("M40");
                    final HSSFRow rowM40 = sheet.getRow(cellReferenceM40.getRow());
                    final HSSFCell cellM40 = rowM40.getCell(cellReferenceM40.getCol());
                    final double resultS40 = ((
                            cellG40.getNumericCellValue() + cellJ40.getNumericCellValue() + cellM40.getNumericCellValue()) / 3)
                            - cellD40.getNumericCellValue();
                    final CellReference cellReferenceD41 = new CellReference("D41");
                    final HSSFRow rowD41 = sheet.getRow(cellReferenceD41.getRow());
                    final HSSFCell cellD41 = rowD41.getCell(cellReferenceD41.getCol());
                    final CellReference cellReferenceG41 = new CellReference("G41");
                    final HSSFRow rowG41 = sheet.getRow(cellReferenceG41.getRow());
                    final HSSFCell cellG41 = rowG41.getCell(cellReferenceG41.getCol());
                    final CellReference cellReferenceJ41 = new CellReference("J41");
                    final HSSFRow rowJ41 = sheet.getRow(cellReferenceJ41.getRow());
                    final HSSFCell cellJ41 = rowJ41.getCell(cellReferenceJ41.getCol());
                    final CellReference cellReferenceM41 = new CellReference("M41");
                    final HSSFRow rowM41 = sheet.getRow(cellReferenceM41.getRow());
                    final HSSFCell cellM41 = rowM41.getCell(cellReferenceM41.getCol());
                    final double resultS41 = ((
                            cellG41.getNumericCellValue() + cellJ41.getNumericCellValue() + cellM41.getNumericCellValue()) / 3)
                            - cellD41.getNumericCellValue();
                    /*-----------------492-----------------*/
                    /*-----------------630-----------------*/
                    final CellReference cellReferenceD42 = new CellReference("D42");
                    final HSSFRow rowD42 = sheet.getRow(cellReferenceD42.getRow());
                    final HSSFCell cellD42 = rowD42.getCell(cellReferenceD42.getCol());
                    final CellReference cellReferenceG42 = new CellReference("G42");
                    final HSSFRow rowG42 = sheet.getRow(cellReferenceG42.getRow());
                    final HSSFCell cellG42 = rowG42.getCell(cellReferenceG42.getCol());
                    final CellReference cellReferenceJ42 = new CellReference("J42");
                    final HSSFRow rowJ42 = sheet.getRow(cellReferenceJ42.getRow());
                    final HSSFCell cellJ42 = rowJ42.getCell(cellReferenceJ42.getCol());
                    final CellReference cellReferenceM42 = new CellReference("M42");
                    final HSSFRow rowM42 = sheet.getRow(cellReferenceM42.getRow());
                    final HSSFCell cellM42 = rowM42.getCell(cellReferenceM42.getCol());
                    final double resultS42 = ((
                            cellG42.getNumericCellValue() + cellJ42.getNumericCellValue() + cellM42.getNumericCellValue()) / 3)
                            - cellD42.getNumericCellValue();
                    final CellReference cellReferenceD43 = new CellReference("D43");
                    final HSSFRow rowD43 = sheet.getRow(cellReferenceD43.getRow());
                    final HSSFCell cellD43 = rowD43.getCell(cellReferenceD43.getCol());
                    final CellReference cellReferenceG43 = new CellReference("G43");
                    final HSSFRow rowG43 = sheet.getRow(cellReferenceG43.getRow());
                    final HSSFCell cellG43 = rowG43.getCell(cellReferenceG43.getCol());
                    final CellReference cellReferenceJ43 = new CellReference("J43");
                    final HSSFRow rowJ43 = sheet.getRow(cellReferenceJ43.getRow());
                    final HSSFCell cellJ43 = rowJ43.getCell(cellReferenceJ43.getCol());
                    final CellReference cellReferenceM43 = new CellReference("M43");
                    final HSSFRow rowM43 = sheet.getRow(cellReferenceM43.getRow());
                    final HSSFCell cellM43 = rowM43.getCell(cellReferenceM43.getCol());
                    final double resultS43 = ((
                            cellG43.getNumericCellValue() + cellJ43.getNumericCellValue() + cellM43.getNumericCellValue()) / 3)
                            - cellD43.getNumericCellValue();
                    final CellReference cellReferenceD44 = new CellReference("D44");
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
                    final ArrayList<Double> valueTolerance = new ArrayList<>();
                    valueTolerance.add(resultS27);
                    valueTolerance.add(resultS28);
                    valueTolerance.add(resultS29);
                    valueTolerance.add(resultS30);
                    valueTolerance.add(resultS31);
                    valueTolerance.add(resultS32);
                    valueTolerance.add(resultS33);
                    valueTolerance.add(resultS34);
                    valueTolerance.add(resultS35);
                    valueTolerance.add(resultS36);
                    valueTolerance.add(resultS37);
                    valueTolerance.add(resultS38);
                    valueTolerance.add(resultS39);
                    valueTolerance.add(resultS40);
                    valueTolerance.add(resultS41);
                    valueTolerance.add(resultS42);
                    valueTolerance.add(resultS43);
                    valueTolerance.add(resultS44);
                    valueTolerance.add(resultS45);
//                            valueTolerance.add(resultS46);
                    final Double maxValueTolerance = Collections.max(valueTolerance);
                    final Double minValueTolerance = Collections.min(valueTolerance);
                    final double absMaxValueTolerance = Math.abs(maxValueTolerance);
                    final double absMinValueTolerance = Math.abs(minValueTolerance);
                    String absMaxValueToleranceString = String.format("%.4f", maxValueTolerance);
                    String absMinValueToleranceString = String.format("%.4f", minValueTolerance);
                    Double showNumber = 0.00;
                    double biggerAbsNumber = 0;
                    if (absMaxValueTolerance > absMinValueTolerance) {
                        biggerAbsNumber = maxValueTolerance;
                        showNumber = absMaxValueTolerance;
                    } else if (absMaxValueTolerance < absMinValueTolerance) {
                        biggerAbsNumber = minValueTolerance;
                        showNumber = absMinValueTolerance;
                    } else if (absMaxValueTolerance == absMinValueTolerance) {
                        biggerAbsNumber = maxValueTolerance;
                        showNumber = absMaxValueTolerance;
                    }
                    String biggerAbsValueToleranceString = String.format("%.4f", biggerAbsNumber);
                    textFieldValueTolerance.setText(biggerAbsValueToleranceString);
                    if (showNumber >= 0.015 || showNumber <= -0.015) {
                        labelValueToleranceResult.setText("错误");
                        labelValueToleranceResult.setForeground(Color.red);
                    } else {
                        labelValueToleranceResult.setText("正确");
                        labelValueToleranceResult.setForeground(Color.green);
                    }
                    /*-----------------630-----------------*/
                    /*-----读取吸光度单元格-----*/
                    /*-----读取吸光度重复值单元格-----*/
                    final CellReference cellReferenceD53 = new CellReference("D53");
                    final HSSFRow rowD53 = sheet.getRow(cellReferenceD53.getRow());
                    final HSSFCell cellD53 = rowD53.getCell(cellReferenceD53.getCol());
                    final CellReference cellReferenceF53 = new CellReference("F53");
                    final HSSFRow rowF53 = sheet.getRow(cellReferenceF53.getRow());
                    final HSSFCell cellF53 = rowF53.getCell(cellReferenceF53.getCol());
                    final CellReference cellReferenceH53 = new CellReference("H53");
                    final HSSFRow rowH53 = sheet.getRow(cellReferenceH53.getRow());
                    final HSSFCell cellH53 = rowH53.getCell(cellReferenceH53.getCol());
                    final CellReference cellReferenceJ53 = new CellReference("J53");
                    final HSSFRow rowJ53 = sheet.getRow(cellReferenceJ53.getRow());
                    final HSSFCell cellJ53 = rowJ53.getCell(cellReferenceJ53.getCol());
                    final CellReference cellReferenceL53 = new CellReference("L53");
                    final HSSFRow rowL53 = sheet.getRow(cellReferenceL53.getRow());
                    final HSSFCell cellL53 = rowL53.getCell(cellReferenceL53.getCol());
                    final CellReference cellReferenceN53 = new CellReference("N53");
                    final HSSFRow rowN53 = sheet.getRow(cellReferenceN53.getRow());
                    final HSSFCell cellN53 = rowN53.getCell(cellReferenceN53.getCol());
                    final CellReference cellReferenceP53 = new CellReference("P53");
                    final HSSFRow rowP53 = sheet.getRow(cellReferenceP53.getRow());
                    final HSSFCell cellP53 = rowP53.getCell(cellReferenceP53.getCol());
                    final double averagep53 =
                            (
                                    cellD53.getNumericCellValue() +
                                            cellF53.getNumericCellValue() +
                                            cellH53.getNumericCellValue() +
                                            cellJ53.getNumericCellValue() +
                                            cellL53.getNumericCellValue() +
                                            cellN53.getNumericCellValue()
                            ) / 6;
                    final double dp53 = Math.pow((cellD53.getNumericCellValue() - averagep53), 2);
                    final double fp53 = Math.pow((cellF53.getNumericCellValue() - averagep53), 2);
                    final double hp53 = Math.pow((cellH53.getNumericCellValue() - averagep53), 2);
                    final double jp53 = Math.pow((cellJ53.getNumericCellValue() - averagep53), 2);
                    final double lp53 = Math.pow((cellL53.getNumericCellValue() - averagep53), 2);
                    final double np53 = Math.pow((cellN53.getNumericCellValue() - averagep53), 2);
                    double resultrsd = (Math.pow((dp53 + fp53 + hp53 + jp53 + lp53 + np53) / 5, 0.5)) / averagep53;
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
                    /*-----读取灵敏度单元格-----*/
                    final CellReference cellReferenceD57 = new CellReference("D57");
                    final HSSFRow rowD57 = sheet.getRow(cellReferenceD57.getRow());
                    final HSSFCell cellD57 = rowD57.getCell(cellReferenceD57.getCol());
                    final double resultSensitive = cellD57.getNumericCellValue();
                    textFieldSensitive.setText(String.valueOf(resultSensitive));
                    if (resultSensitive > 0.05) {
                        labelSensitiveResult.setText("正确");
                        labelSensitiveResult.setForeground(Color.green);
                    } else {
                        labelSensitiveResult.setText("错误");
                        labelSensitiveResult.setForeground(Color.red);
                    }
                    /*-----读取灵敏度单元格-----*/
                    /*-----读取通道差异单元格-----*/
                    /*-----------------TunnalA-----------------*/
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
                    final CellReference cellReferenceP61 = new CellReference("P61");
                    final HSSFRow rowP61 = sheet.getRow(cellReferenceP61.getRow());
                    final HSSFCell cellP61 = rowP61.getCell(cellReferenceP61.getCol());
                    final CellReference cellReferenceS61 = new CellReference("S61");
                    final HSSFRow rowS61 = sheet.getRow(cellReferenceS61.getRow());
                    final HSSFCell cellS61 = rowS61.getCell(cellReferenceS61.getCol());
                    final CellReference cellReferenceV61 = new CellReference("V61");
                    final HSSFRow rowV61 = sheet.getRow(cellReferenceV61.getRow());
                    final HSSFCell cellV61 = rowV61.getCell(cellReferenceV61.getCol());
                    final CellReference cellReferenceZ61 = new CellReference("Z61");
                    final HSSFRow rowZ61 = sheet.getRow(cellReferenceZ61.getRow());
                    final HSSFCell cellZ61 = rowZ61.getCell(cellReferenceZ61.getCol());
                    final ArrayList<Double> tunnalA = new ArrayList<>();
                    tunnalA.add(cellD61.getNumericCellValue());
                    tunnalA.add(cellG61.getNumericCellValue());
                    tunnalA.add(cellJ61.getNumericCellValue());
                    tunnalA.add(cellM61.getNumericCellValue());
                    tunnalA.add(cellP61.getNumericCellValue());
                    tunnalA.add(cellS61.getNumericCellValue());
                    double maxTunnalA = Collections.max(tunnalA);
                    double minTunnalA = Collections.min(tunnalA);
                    maxTunnalA = cellV61.getNumericCellValue();
                    minTunnalA = cellZ61.getNumericCellValue();
                    /*-----------------TunnalA-----------------*/
                    /*-----------------TunnalB-----------------*/
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
                    final CellReference cellReferenceP62 = new CellReference("P62");
                    final HSSFRow rowP62 = sheet.getRow(cellReferenceP62.getRow());
                    final HSSFCell cellP62 = rowP62.getCell(cellReferenceP62.getCol());
                    final CellReference cellReferenceS62 = new CellReference("S62");
                    final HSSFRow rowS62 = sheet.getRow(cellReferenceS62.getRow());
                    final HSSFCell cellS62 = rowS62.getCell(cellReferenceS62.getCol());
                    final CellReference cellReferenceV62 = new CellReference("V62");
                    final HSSFRow rowV62 = sheet.getRow(cellReferenceV62.getRow());
                    final HSSFCell cellV62 = rowV62.getCell(cellReferenceV62.getCol());
                    final CellReference cellReferenceZ62 = new CellReference("Z62");
                    final HSSFRow rowZ62 = sheet.getRow(cellReferenceZ62.getRow());
                    final HSSFCell cellZ62 = rowZ62.getCell(cellReferenceZ62.getCol());
                    final ArrayList<Double> tunnalB = new ArrayList<>();
                    tunnalB.add(cellD62.getNumericCellValue());
                    tunnalB.add(cellG62.getNumericCellValue());
                    tunnalB.add(cellJ62.getNumericCellValue());
                    tunnalB.add(cellM62.getNumericCellValue());
                    tunnalB.add(cellP62.getNumericCellValue());
                    tunnalB.add(cellS62.getNumericCellValue());
                    double maxtunnalB = Collections.max(tunnalB);
                    double mintunnalB = Collections.min(tunnalB);
                    maxtunnalB = cellV62.getNumericCellValue();
                    mintunnalB = cellZ62.getNumericCellValue();
                    /*-----------------TunnalB-----------------*/
                    /*-----------------TunnalC-----------------*/
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
                    final CellReference cellReferenceP63 = new CellReference("P63");
                    final HSSFRow rowP63 = sheet.getRow(cellReferenceP63.getRow());
                    final HSSFCell cellP63 = rowP63.getCell(cellReferenceP63.getCol());
                    final CellReference cellReferenceS63 = new CellReference("S63");
                    final HSSFRow rowS63 = sheet.getRow(cellReferenceS63.getRow());
                    final HSSFCell cellS63 = rowS63.getCell(cellReferenceS63.getCol());
                    final CellReference cellReferenceV63 = new CellReference("V63");
                    final HSSFRow rowV63 = sheet.getRow(cellReferenceV63.getRow());
                    final HSSFCell cellV63 = rowV63.getCell(cellReferenceV63.getCol());
                    final CellReference cellReferenceZ63 = new CellReference("Z63");
                    final HSSFRow rowZ63 = sheet.getRow(cellReferenceZ63.getRow());
                    final HSSFCell cellZ63 = rowZ63.getCell(cellReferenceZ63.getCol());
                    final ArrayList<Double> tunnalC = new ArrayList<>();
                    tunnalC.add(cellD63.getNumericCellValue());
                    tunnalC.add(cellG63.getNumericCellValue());
                    tunnalC.add(cellJ63.getNumericCellValue());
                    tunnalC.add(cellM63.getNumericCellValue());
                    tunnalC.add(cellP63.getNumericCellValue());
                    tunnalC.add(cellS63.getNumericCellValue());
                    double maxtunnalC = Collections.max(tunnalC);
                    double mintunnalC = Collections.min(tunnalC);
                    maxtunnalC = cellV63.getNumericCellValue();
                    mintunnalC = cellZ63.getNumericCellValue();
                    /*-----------------TunnalC-----------------*/
                    /*-----------------TunnalD-----------------*/
                    final CellReference cellReferenceD64 = new CellReference("D64");
                    final HSSFRow rowD64 = sheet.getRow(cellReferenceD64.getRow());
                    final HSSFCell cellD64 = rowD64.getCell(cellReferenceD64.getCol());
                    final CellReference cellReferenceG64 = new CellReference("G64");
                    final HSSFRow rowG64 = sheet.getRow(cellReferenceG64.getRow());
                    final HSSFCell cellG64 = rowG64.getCell(cellReferenceG64.getCol());
                    final CellReference cellReferenceJ64 = new CellReference("J64");
                    final HSSFRow rowJ64 = sheet.getRow(cellReferenceJ64.getRow());
                    final HSSFCell cellJ64 = rowJ64.getCell(cellReferenceJ64.getCol());
                    final CellReference cellReferenceM64 = new CellReference("M64");
                    final HSSFRow rowM64 = sheet.getRow(cellReferenceM64.getRow());
                    final HSSFCell cellM64 = rowM64.getCell(cellReferenceM64.getCol());
                    final CellReference cellReferenceP64 = new CellReference("P64");
                    final HSSFRow rowP64 = sheet.getRow(cellReferenceP64.getRow());
                    final HSSFCell cellP64 = rowP64.getCell(cellReferenceP64.getCol());
                    final CellReference cellReferenceS64 = new CellReference("S64");
                    final HSSFRow rowS64 = sheet.getRow(cellReferenceS64.getRow());
                    final HSSFCell cellS64 = rowS64.getCell(cellReferenceS64.getCol());
                    final CellReference cellReferenceV64 = new CellReference("V64");
                    final HSSFRow rowV64 = sheet.getRow(cellReferenceV64.getRow());
                    final HSSFCell cellV64 = rowV64.getCell(cellReferenceV64.getCol());
                    final CellReference cellReferenceZ64 = new CellReference("Z64");
                    final HSSFRow rowZ64 = sheet.getRow(cellReferenceZ64.getRow());
                    final HSSFCell cellZ64 = rowZ64.getCell(cellReferenceZ64.getCol());
                    final ArrayList<Double> tunnalD = new ArrayList<>();
                    tunnalD.add(cellD64.getNumericCellValue());
                    tunnalD.add(cellG64.getNumericCellValue());
                    tunnalD.add(cellJ64.getNumericCellValue());
                    tunnalD.add(cellM64.getNumericCellValue());
                    tunnalD.add(cellP64.getNumericCellValue());
                    tunnalD.add(cellS64.getNumericCellValue());
                    double maxtunnalD = Collections.max(tunnalD);
                    double mintunnalD = Collections.min(tunnalD);
                    maxtunnalD = cellV64.getNumericCellValue();
                    mintunnalD = cellZ64.getNumericCellValue();
                    /*-----------------TunnalD-----------------*/
                    /*-----------------TunnalE-----------------*/
                    final CellReference cellReferenceD65 = new CellReference("D65");
                    final HSSFRow rowD65 = sheet.getRow(cellReferenceD65.getRow());
                    final HSSFCell cellD65 = rowD65.getCell(cellReferenceD65.getCol());
                    final CellReference cellReferenceG65 = new CellReference("G65");
                    final HSSFRow rowG65 = sheet.getRow(cellReferenceG65.getRow());
                    final HSSFCell cellG65 = rowG65.getCell(cellReferenceG65.getCol());
                    final CellReference cellReferenceJ65 = new CellReference("J65");
                    final HSSFRow rowJ65 = sheet.getRow(cellReferenceJ65.getRow());
                    final HSSFCell cellJ65 = rowJ65.getCell(cellReferenceJ65.getCol());
                    final CellReference cellReferenceM65 = new CellReference("M65");
                    final HSSFRow rowM65 = sheet.getRow(cellReferenceM65.getRow());
                    final HSSFCell cellM65 = rowM65.getCell(cellReferenceM65.getCol());
                    final CellReference cellReferenceP65 = new CellReference("P65");
                    final HSSFRow rowP65 = sheet.getRow(cellReferenceP65.getRow());
                    final HSSFCell cellP65 = rowP65.getCell(cellReferenceP65.getCol());
                    final CellReference cellReferenceS65 = new CellReference("S65");
                    final HSSFRow rowS65 = sheet.getRow(cellReferenceS65.getRow());
                    final HSSFCell cellS65 = rowS65.getCell(cellReferenceS65.getCol());
                    final CellReference cellReferenceV65 = new CellReference("V65");
                    final HSSFRow rowV65 = sheet.getRow(cellReferenceV65.getRow());
                    final HSSFCell cellV65 = rowV65.getCell(cellReferenceV65.getCol());
                    final CellReference cellReferenceZ65 = new CellReference("Z65");
                    final HSSFRow rowZ65 = sheet.getRow(cellReferenceZ65.getRow());
                    final HSSFCell cellZ65 = rowZ65.getCell(cellReferenceZ65.getCol());
                    final ArrayList<Double> tunnalE = new ArrayList<>();
                    tunnalE.add(cellD65.getNumericCellValue());
                    tunnalE.add(cellG65.getNumericCellValue());
                    tunnalE.add(cellJ65.getNumericCellValue());
                    tunnalE.add(cellM65.getNumericCellValue());
                    tunnalE.add(cellP65.getNumericCellValue());
                    tunnalE.add(cellS65.getNumericCellValue());
                    double maxtunnalE = Collections.max(tunnalE);
                    double mintunnalE = Collections.min(tunnalE);
                    maxtunnalE = cellV65.getNumericCellValue();
                    mintunnalE = cellZ65.getNumericCellValue();
                    /*-----------------TunnalE-----------------*/
                    /*-----------------TunnalF-----------------*/
                    final CellReference cellReferenceD66 = new CellReference("D66");
                    final HSSFRow rowD66 = sheet.getRow(cellReferenceD66.getRow());
                    final HSSFCell cellD66 = rowD66.getCell(cellReferenceD66.getCol());
                    final CellReference cellReferenceG66 = new CellReference("G66");
                    final HSSFRow rowG66 = sheet.getRow(cellReferenceG66.getRow());
                    final HSSFCell cellG66 = rowG66.getCell(cellReferenceG66.getCol());
                    final CellReference cellReferenceJ66 = new CellReference("J66");
                    final HSSFRow rowJ66 = sheet.getRow(cellReferenceJ66.getRow());
                    final HSSFCell cellJ66 = rowJ66.getCell(cellReferenceJ66.getCol());
                    final CellReference cellReferenceM66 = new CellReference("M66");
                    final HSSFRow rowM66 = sheet.getRow(cellReferenceM66.getRow());
                    final HSSFCell cellM66 = rowM66.getCell(cellReferenceM66.getCol());
                    final CellReference cellReferenceP66 = new CellReference("P66");
                    final HSSFRow rowP66 = sheet.getRow(cellReferenceP66.getRow());
                    final HSSFCell cellP66 = rowP66.getCell(cellReferenceP66.getCol());
                    final CellReference cellReferenceS66 = new CellReference("S66");
                    final HSSFRow rowS66 = sheet.getRow(cellReferenceS66.getRow());
                    final HSSFCell cellS66 = rowS66.getCell(cellReferenceS66.getCol());
                    final CellReference cellReferenceV66 = new CellReference("V66");
                    final HSSFRow rowV66 = sheet.getRow(cellReferenceV66.getRow());
                    final HSSFCell cellV66 = rowV66.getCell(cellReferenceV66.getCol());
                    final CellReference cellReferenceZ66 = new CellReference("Z66");
                    final HSSFRow rowZ66 = sheet.getRow(cellReferenceZ66.getRow());
                    final HSSFCell cellZ66 = rowZ66.getCell(cellReferenceZ66.getCol());
                    final ArrayList<Double> tunnalF = new ArrayList<>();
                    tunnalF.add(cellD66.getNumericCellValue());
                    tunnalF.add(cellG66.getNumericCellValue());
                    tunnalF.add(cellJ66.getNumericCellValue());
                    tunnalF.add(cellM66.getNumericCellValue());
                    tunnalF.add(cellP66.getNumericCellValue());
                    tunnalF.add(cellS66.getNumericCellValue());
                    double maxtunnalF = Collections.max(tunnalF);
                    double mintunnalF = Collections.min(tunnalF);
                    maxtunnalF = cellV66.getNumericCellValue();
                    mintunnalF = cellZ66.getNumericCellValue();
                    /*-----------------TunnalF-----------------*/
                    /*-----------------TunnalG-----------------*/
                    final CellReference cellReferenceD67 = new CellReference("D67");
                    final HSSFRow rowD67 = sheet.getRow(cellReferenceD67.getRow());
                    final HSSFCell cellD67 = rowD67.getCell(cellReferenceD67.getCol());
                    final CellReference cellReferenceG67 = new CellReference("G67");
                    final HSSFRow rowG67 = sheet.getRow(cellReferenceG67.getRow());
                    final HSSFCell cellG67 = rowG67.getCell(cellReferenceG67.getCol());
                    final CellReference cellReferenceJ67 = new CellReference("J67");
                    final HSSFRow rowJ67 = sheet.getRow(cellReferenceJ67.getRow());
                    final HSSFCell cellJ67 = rowJ67.getCell(cellReferenceJ67.getCol());
                    final CellReference cellReferenceM67 = new CellReference("M67");
                    final HSSFRow rowM67 = sheet.getRow(cellReferenceM67.getRow());
                    final HSSFCell cellM67 = rowM67.getCell(cellReferenceM67.getCol());
                    final CellReference cellReferenceP67 = new CellReference("P67");
                    final HSSFRow rowP67 = sheet.getRow(cellReferenceP67.getRow());
                    final HSSFCell cellP67 = rowP67.getCell(cellReferenceP67.getCol());
                    final CellReference cellReferenceS67 = new CellReference("S67");
                    final HSSFRow rowS67 = sheet.getRow(cellReferenceS67.getRow());
                    final HSSFCell cellS67 = rowS67.getCell(cellReferenceS67.getCol());
                    final CellReference cellReferenceV67 = new CellReference("V67");
                    final HSSFRow rowV67 = sheet.getRow(cellReferenceV67.getRow());
                    final HSSFCell cellV67 = rowV67.getCell(cellReferenceV67.getCol());
                    final CellReference cellReferenceZ67 = new CellReference("Z67");
                    final HSSFRow rowZ67 = sheet.getRow(cellReferenceZ67.getRow());
                    final HSSFCell cellZ67 = rowZ67.getCell(cellReferenceZ67.getCol());
                    final ArrayList<Double> tunnalG = new ArrayList<>();
                    tunnalG.add(cellD67.getNumericCellValue());
                    tunnalG.add(cellG67.getNumericCellValue());
                    tunnalG.add(cellJ67.getNumericCellValue());
                    tunnalG.add(cellM67.getNumericCellValue());
                    tunnalG.add(cellP67.getNumericCellValue());
                    tunnalG.add(cellS67.getNumericCellValue());
                    double maxtunnalG = Collections.max(tunnalG);
                    double mintunnalG = Collections.min(tunnalG);
                    maxtunnalG = cellV67.getNumericCellValue();
                    mintunnalG = cellZ67.getNumericCellValue();
                    /*-----------------TunnalG-----------------*/
                    /*-----------------TunnalH-----------------*/
                    final CellReference cellReferenceD68 = new CellReference("D68");
                    final HSSFRow rowD68 = sheet.getRow(cellReferenceD68.getRow());
                    final HSSFCell cellD68 = rowD68.getCell(cellReferenceD68.getCol());
                    final CellReference cellReferenceG68 = new CellReference("G68");
                    final HSSFRow rowG68 = sheet.getRow(cellReferenceG68.getRow());
                    final HSSFCell cellG68 = rowG68.getCell(cellReferenceG68.getCol());
                    final CellReference cellReferenceJ68 = new CellReference("J68");
                    final HSSFRow rowJ68 = sheet.getRow(cellReferenceJ68.getRow());
                    final HSSFCell cellJ68 = rowJ68.getCell(cellReferenceJ68.getCol());
                    final CellReference cellReferenceM68 = new CellReference("M68");
                    final HSSFRow rowM68 = sheet.getRow(cellReferenceM68.getRow());
                    final HSSFCell cellM68 = rowM68.getCell(cellReferenceM68.getCol());
                    final CellReference cellReferenceP68 = new CellReference("P68");
                    final HSSFRow rowP68 = sheet.getRow(cellReferenceP68.getRow());
                    final HSSFCell cellP68 = rowP68.getCell(cellReferenceP68.getCol());
                    final CellReference cellReferenceS68 = new CellReference("S68");
                    final HSSFRow rowS68 = sheet.getRow(cellReferenceS68.getRow());
                    final HSSFCell cellS68 = rowS68.getCell(cellReferenceS68.getCol());
                    final CellReference cellReferenceV68 = new CellReference("V68");
                    final HSSFRow rowV68 = sheet.getRow(cellReferenceV68.getRow());
                    final HSSFCell cellV68 = rowV68.getCell(cellReferenceV68.getCol());
                    final CellReference cellReferenceZ68 = new CellReference("Z68");
                    final HSSFRow rowZ68 = sheet.getRow(cellReferenceZ68.getRow());
                    final HSSFCell cellZ68 = rowZ68.getCell(cellReferenceZ68.getCol());
                    final ArrayList<Double> tunnalH = new ArrayList<>();
                    tunnalH.add(cellD68.getNumericCellValue());
                    tunnalH.add(cellG68.getNumericCellValue());
                    tunnalH.add(cellJ68.getNumericCellValue());
                    tunnalH.add(cellM68.getNumericCellValue());
                    tunnalH.add(cellP68.getNumericCellValue());
                    tunnalH.add(cellS68.getNumericCellValue());
                    double maxtunnalH = Collections.max(tunnalH);
                    double mintunnalH = Collections.min(tunnalH);
                    maxtunnalH = cellV68.getNumericCellValue();
                    mintunnalH = cellZ68.getNumericCellValue();
                    /*-----------------TunnalH-----------------*/
                    /*-----------------Amax-----------------*/
                    final ArrayList<Double> Amax = new ArrayList<>();
                    Amax.add(maxTunnalA);
                    Amax.add(maxtunnalB);
                    Amax.add(maxtunnalC);
                    Amax.add(maxtunnalD);
                    Amax.add(maxtunnalE);
                    Amax.add(maxtunnalF);
                    Amax.add(maxtunnalG);
                    Amax.add(maxtunnalH);
                    double resultAmax = Collections.max(Amax);
                    /*-----------------Amax-----------------*/
                    /*-----------------Amin-----------------*/
                    final ArrayList<Double> Amin = new ArrayList<>();
                    Amin.add(minTunnalA);
                    Amin.add(mintunnalB);
                    Amin.add(mintunnalC);
                    Amin.add(mintunnalD);
                    Amin.add(mintunnalE);
                    Amin.add(mintunnalF);
                    Amin.add(mintunnalG);
                    Amin.add(mintunnalH);
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
                    /*-----读取机身号单元格数据-----*/
                    final CellReference cellReferenceP5 = new CellReference("P5");
                    final HSSFRow rowP5 = sheet.getRow(cellReferenceP5.getRow());
                    final HSSFCell cellP5 = rowP5.getCell(cellReferenceP5.getCol());
                    String cnProductNumber = cellP5.getStringCellValue();
                    final Matcher cnmatcherProductNumber = pattern.matcher(cnProductNumber.replaceAll("\\s", ""));
                    /*-----读取机身号单元格数据-----*/
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
                        /*-----读取灵敏度单元格-----*/
                        final CellReference cellReferenceD74 = new CellReference("D74");
                        final HSSFRow rowD74 = sheet.getRow(cellReferenceD74.getRow());
                        final HSSFCell cellD74 = rowD74.getCell(cellReferenceD74.getCol());
                        final double resultSensitive = cellD74.getNumericCellValue();
                        textFieldSensitive.setText(String.valueOf(resultSensitive));
                        if (resultSensitive > 0.05) {
                            labelSensitiveResult.setText("正确");
                            labelSensitiveResult.setForeground(Color.green);
                        } else {
                            labelSensitiveResult.setText("错误");
                            labelSensitiveResult.setForeground(Color.red);
                        }
                        /*-----读取灵敏度单元格-----*/
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
                        /*-----------------TunnalC-----------------*/
                        final CellReference cellReferenceD80 = new CellReference("D80");
                        final HSSFRow rowD80 = sheet.getRow(cellReferenceD80.getRow());
                        final HSSFCell cellD80 = rowD80.getCell(cellReferenceD80.getCol());
                        final CellReference cellReferenceG80 = new CellReference("G80");
                        final HSSFRow rowG80 = sheet.getRow(cellReferenceG80.getRow());
                        final HSSFCell cellG80 = rowG80.getCell(cellReferenceG80.getCol());
                        final CellReference cellReferenceJ80 = new CellReference("J80");
                        final HSSFRow rowJ80 = sheet.getRow(cellReferenceJ80.getRow());
                        final HSSFCell cellJ80 = rowJ80.getCell(cellReferenceJ80.getCol());
                        final CellReference cellReferenceM80 = new CellReference("M80");
                        final HSSFRow rowM80 = sheet.getRow(cellReferenceM80.getRow());
                        final HSSFCell cellM80 = rowM80.getCell(cellReferenceM80.getCol());
                        final CellReference cellReferenceP80 = new CellReference("P80");
                        final HSSFRow rowP80 = sheet.getRow(cellReferenceP80.getRow());
                        final HSSFCell cellP80 = rowP80.getCell(cellReferenceP80.getCol());
                        final CellReference cellReferenceS80 = new CellReference("S80");
                        final HSSFRow rowS80 = sheet.getRow(cellReferenceS80.getRow());
                        final HSSFCell cellS80 = rowS80.getCell(cellReferenceS80.getCol());
                        final CellReference cellReferenceV80 = new CellReference("V80");
                        final HSSFRow rowV80 = sheet.getRow(cellReferenceV80.getRow());
                        final HSSFCell cellV80 = rowV80.getCell(cellReferenceV80.getCol());
                        final CellReference cellReferenceZ80 = new CellReference("Z80");
                        final HSSFRow rowZ80 = sheet.getRow(cellReferenceZ80.getRow());
                        final HSSFCell cellZ80 = rowZ80.getCell(cellReferenceZ80.getCol());
                        final ArrayList<Double> tunnalC = new ArrayList<>();
                        tunnalC.add(cellD80.getNumericCellValue());
                        tunnalC.add(cellG80.getNumericCellValue());
                        tunnalC.add(cellJ80.getNumericCellValue());
                        tunnalC.add(cellM80.getNumericCellValue());
                        tunnalC.add(cellP80.getNumericCellValue());
                        tunnalC.add(cellS80.getNumericCellValue());
                        double maxtunnalC = Collections.max(tunnalC);
                        double mintunnalC = Collections.min(tunnalC);
                        maxtunnalC = cellV80.getNumericCellValue();
                        mintunnalC = cellZ80.getNumericCellValue();
                        /*-----------------TunnalC-----------------*/
                        /*-----------------TunnalD-----------------*/
                        final CellReference cellReferenceD81 = new CellReference("D81");
                        final HSSFRow rowD81 = sheet.getRow(cellReferenceD81.getRow());
                        final HSSFCell cellD81 = rowD81.getCell(cellReferenceD81.getCol());
                        final CellReference cellReferenceG81 = new CellReference("G81");
                        final HSSFRow rowG81 = sheet.getRow(cellReferenceG81.getRow());
                        final HSSFCell cellG81 = rowG81.getCell(cellReferenceG81.getCol());
                        final CellReference cellReferenceJ81 = new CellReference("J81");
                        final HSSFRow rowJ81 = sheet.getRow(cellReferenceJ81.getRow());
                        final HSSFCell cellJ81 = rowJ81.getCell(cellReferenceJ81.getCol());
                        final CellReference cellReferenceM81 = new CellReference("M81");
                        final HSSFRow rowM81 = sheet.getRow(cellReferenceM81.getRow());
                        final HSSFCell cellM81 = rowM81.getCell(cellReferenceM81.getCol());
                        final CellReference cellReferenceP81 = new CellReference("P81");
                        final HSSFRow rowP81 = sheet.getRow(cellReferenceP81.getRow());
                        final HSSFCell cellP81 = rowP81.getCell(cellReferenceP81.getCol());
                        final CellReference cellReferenceS81 = new CellReference("S81");
                        final HSSFRow rowS81 = sheet.getRow(cellReferenceS81.getRow());
                        final HSSFCell cellS81 = rowS81.getCell(cellReferenceS81.getCol());
                        final CellReference cellReferenceV81 = new CellReference("V81");
                        final HSSFRow rowV81 = sheet.getRow(cellReferenceV81.getRow());
                        final HSSFCell cellV81 = rowV81.getCell(cellReferenceV81.getCol());
                        final CellReference cellReferenceZ81 = new CellReference("Z81");
                        final HSSFRow rowZ81 = sheet.getRow(cellReferenceZ81.getRow());
                        final HSSFCell cellZ81 = rowZ81.getCell(cellReferenceZ81.getCol());
                        final ArrayList<Double> tunnalD = new ArrayList<>();
                        tunnalD.add(cellD81.getNumericCellValue());
                        tunnalD.add(cellG81.getNumericCellValue());
                        tunnalD.add(cellJ81.getNumericCellValue());
                        tunnalD.add(cellM81.getNumericCellValue());
                        tunnalD.add(cellP81.getNumericCellValue());
                        tunnalD.add(cellS81.getNumericCellValue());
                        double maxtunnalD = Collections.max(tunnalD);
                        double mintunnalD = Collections.min(tunnalD);
                        maxtunnalD = cellV81.getNumericCellValue();
                        mintunnalD = cellZ81.getNumericCellValue();
                        /*-----------------TunnalD-----------------*/
                        /*-----------------TunnalE-----------------*/
                        final CellReference cellReferenceD82 = new CellReference("D82");
                        final HSSFRow rowD82 = sheet.getRow(cellReferenceD82.getRow());
                        final HSSFCell cellD82 = rowD82.getCell(cellReferenceD82.getCol());
                        final CellReference cellReferenceG82 = new CellReference("G82");
                        final HSSFRow rowG82 = sheet.getRow(cellReferenceG82.getRow());
                        final HSSFCell cellG82 = rowG82.getCell(cellReferenceG82.getCol());
                        final CellReference cellReferenceJ82 = new CellReference("J82");
                        final HSSFRow rowJ82 = sheet.getRow(cellReferenceJ82.getRow());
                        final HSSFCell cellJ82 = rowJ82.getCell(cellReferenceJ82.getCol());
                        final CellReference cellReferenceM82 = new CellReference("M82");
                        final HSSFRow rowM82 = sheet.getRow(cellReferenceM82.getRow());
                        final HSSFCell cellM82 = rowM82.getCell(cellReferenceM82.getCol());
                        final CellReference cellReferenceP82 = new CellReference("P82");
                        final HSSFRow rowP82 = sheet.getRow(cellReferenceP82.getRow());
                        final HSSFCell cellP82 = rowP82.getCell(cellReferenceP82.getCol());
                        final CellReference cellReferenceS82 = new CellReference("S82");
                        final HSSFRow rowS82 = sheet.getRow(cellReferenceS82.getRow());
                        final HSSFCell cellS82 = rowS82.getCell(cellReferenceS82.getCol());
                        final CellReference cellReferenceV82 = new CellReference("V82");
                        final HSSFRow rowV82 = sheet.getRow(cellReferenceV82.getRow());
                        final HSSFCell cellV82 = rowV82.getCell(cellReferenceV82.getCol());
                        final CellReference cellReferenceZ82 = new CellReference("Z82");
                        final HSSFRow rowZ82 = sheet.getRow(cellReferenceZ82.getRow());
                        final HSSFCell cellZ82 = rowZ82.getCell(cellReferenceZ82.getCol());
                        final ArrayList<Double> tunnalE = new ArrayList<>();
                        tunnalE.add(cellD82.getNumericCellValue());
                        tunnalE.add(cellG82.getNumericCellValue());
                        tunnalE.add(cellJ82.getNumericCellValue());
                        tunnalE.add(cellM82.getNumericCellValue());
                        tunnalE.add(cellP82.getNumericCellValue());
                        tunnalE.add(cellS82.getNumericCellValue());
                        double maxtunnalE = Collections.max(tunnalE);
                        double mintunnalE = Collections.min(tunnalE);
                        maxtunnalE = cellV82.getNumericCellValue();
                        mintunnalE = cellZ82.getNumericCellValue();
                        /*-----------------TunnalE-----------------*/
                        /*-----------------TunnalF-----------------*/
                        final CellReference cellReferenceD83 = new CellReference("D83");
                        final HSSFRow rowD83 = sheet.getRow(cellReferenceD83.getRow());
                        final HSSFCell cellD83 = rowD83.getCell(cellReferenceD83.getCol());
                        final CellReference cellReferenceG83 = new CellReference("G83");
                        final HSSFRow rowG83 = sheet.getRow(cellReferenceG83.getRow());
                        final HSSFCell cellG83 = rowG83.getCell(cellReferenceG83.getCol());
                        final CellReference cellReferenceJ83 = new CellReference("J83");
                        final HSSFRow rowJ83 = sheet.getRow(cellReferenceJ83.getRow());
                        final HSSFCell cellJ83 = rowJ83.getCell(cellReferenceJ83.getCol());
                        final CellReference cellReferenceM83 = new CellReference("M83");
                        final HSSFRow rowM83 = sheet.getRow(cellReferenceM83.getRow());
                        final HSSFCell cellM83 = rowM83.getCell(cellReferenceM83.getCol());
                        final CellReference cellReferenceP83 = new CellReference("P83");
                        final HSSFRow rowP83 = sheet.getRow(cellReferenceP83.getRow());
                        final HSSFCell cellP83 = rowP83.getCell(cellReferenceP83.getCol());
                        final CellReference cellReferenceS83 = new CellReference("S83");
                        final HSSFRow rowS83 = sheet.getRow(cellReferenceS83.getRow());
                        final HSSFCell cellS83 = rowS83.getCell(cellReferenceS83.getCol());
                        final CellReference cellReferenceV83 = new CellReference("V83");
                        final HSSFRow rowV83 = sheet.getRow(cellReferenceV83.getRow());
                        final HSSFCell cellV83 = rowV83.getCell(cellReferenceV83.getCol());
                        final CellReference cellReferenceZ83 = new CellReference("Z83");
                        final HSSFRow rowZ83 = sheet.getRow(cellReferenceZ83.getRow());
                        final HSSFCell cellZ83 = rowZ83.getCell(cellReferenceZ83.getCol());
                        final ArrayList<Double> tunnalF = new ArrayList<>();
                        tunnalF.add(cellD83.getNumericCellValue());
                        tunnalF.add(cellG83.getNumericCellValue());
                        tunnalF.add(cellJ83.getNumericCellValue());
                        tunnalF.add(cellM83.getNumericCellValue());
                        tunnalF.add(cellP83.getNumericCellValue());
                        tunnalF.add(cellS83.getNumericCellValue());
                        double maxtunnalF = Collections.max(tunnalF);
                        double mintunnalF = Collections.min(tunnalF);
                        maxtunnalF = cellV83.getNumericCellValue();
                        mintunnalF = cellZ83.getNumericCellValue();
                        /*-----------------TunnalF-----------------*/
                        /*-----------------TunnalG-----------------*/
                        final CellReference cellReferenceD84 = new CellReference("D84");
                        final HSSFRow rowD84 = sheet.getRow(cellReferenceD84.getRow());
                        final HSSFCell cellD84 = rowD84.getCell(cellReferenceD84.getCol());
                        final CellReference cellReferenceG84 = new CellReference("G84");
                        final HSSFRow rowG84 = sheet.getRow(cellReferenceG84.getRow());
                        final HSSFCell cellG84 = rowG84.getCell(cellReferenceG84.getCol());
                        final CellReference cellReferenceJ84 = new CellReference("J84");
                        final HSSFRow rowJ84 = sheet.getRow(cellReferenceJ84.getRow());
                        final HSSFCell cellJ84 = rowJ84.getCell(cellReferenceJ84.getCol());
                        final CellReference cellReferenceM84 = new CellReference("M84");
                        final HSSFRow rowM84 = sheet.getRow(cellReferenceM84.getRow());
                        final HSSFCell cellM84 = rowM84.getCell(cellReferenceM84.getCol());
                        final CellReference cellReferenceP84 = new CellReference("P84");
                        final HSSFRow rowP84 = sheet.getRow(cellReferenceP84.getRow());
                        final HSSFCell cellP84 = rowP84.getCell(cellReferenceP84.getCol());
                        final CellReference cellReferenceS84 = new CellReference("S84");
                        final HSSFRow rowS84 = sheet.getRow(cellReferenceS84.getRow());
                        final HSSFCell cellS84 = rowS84.getCell(cellReferenceS84.getCol());
                        final CellReference cellReferenceV84 = new CellReference("V84");
                        final HSSFRow rowV84 = sheet.getRow(cellReferenceV84.getRow());
                        final HSSFCell cellV84 = rowV84.getCell(cellReferenceV84.getCol());
                        final CellReference cellReferenceZ84 = new CellReference("Z84");
                        final HSSFRow rowZ84 = sheet.getRow(cellReferenceZ84.getRow());
                        final HSSFCell cellZ84 = rowZ84.getCell(cellReferenceZ84.getCol());
                        final ArrayList<Double> tunnalG = new ArrayList<>();
                        tunnalG.add(cellD84.getNumericCellValue());
                        tunnalG.add(cellG84.getNumericCellValue());
                        tunnalG.add(cellJ84.getNumericCellValue());
                        tunnalG.add(cellM84.getNumericCellValue());
                        tunnalG.add(cellP84.getNumericCellValue());
                        tunnalG.add(cellS84.getNumericCellValue());
                        double maxtunnalG = Collections.max(tunnalG);
                        double mintunnalG = Collections.min(tunnalG);
                        maxtunnalG = cellV84.getNumericCellValue();
                        mintunnalG = cellZ84.getNumericCellValue();
                        /*-----------------TunnalG-----------------*/
                        /*-----------------TunnalH-----------------*/
                        final CellReference cellReferenceD85 = new CellReference("D85");
                        final HSSFRow rowD85 = sheet.getRow(cellReferenceD85.getRow());
                        final HSSFCell cellD85 = rowD85.getCell(cellReferenceD85.getCol());
                        final CellReference cellReferenceG85 = new CellReference("G85");
                        final HSSFRow rowG85 = sheet.getRow(cellReferenceG85.getRow());
                        final HSSFCell cellG85 = rowG85.getCell(cellReferenceG85.getCol());
                        final CellReference cellReferenceJ85 = new CellReference("J85");
                        final HSSFRow rowJ85 = sheet.getRow(cellReferenceJ85.getRow());
                        final HSSFCell cellJ85 = rowJ85.getCell(cellReferenceJ85.getCol());
                        final CellReference cellReferenceM85 = new CellReference("M85");
                        final HSSFRow rowM85 = sheet.getRow(cellReferenceM85.getRow());
                        final HSSFCell cellM85 = rowM85.getCell(cellReferenceM85.getCol());
                        final CellReference cellReferenceP85 = new CellReference("P85");
                        final HSSFRow rowP85 = sheet.getRow(cellReferenceP85.getRow());
                        final HSSFCell cellP85 = rowP85.getCell(cellReferenceP85.getCol());
                        final CellReference cellReferenceS85 = new CellReference("S85");
                        final HSSFRow rowS85 = sheet.getRow(cellReferenceS85.getRow());
                        final HSSFCell cellS85 = rowS85.getCell(cellReferenceS85.getCol());
                        final CellReference cellReferenceV85 = new CellReference("V85");
                        final HSSFRow rowV85 = sheet.getRow(cellReferenceV85.getRow());
                        final HSSFCell cellV85 = rowV85.getCell(cellReferenceV85.getCol());
                        final CellReference cellReferenceZ85 = new CellReference("Z85");
                        final HSSFRow rowZ85 = sheet.getRow(cellReferenceZ85.getRow());
                        final HSSFCell cellZ85 = rowZ85.getCell(cellReferenceZ85.getCol());
                        final ArrayList<Double> tunnalH = new ArrayList<>();
                        tunnalH.add(cellD85.getNumericCellValue());
                        tunnalH.add(cellG85.getNumericCellValue());
                        tunnalH.add(cellJ85.getNumericCellValue());
                        tunnalH.add(cellM85.getNumericCellValue());
                        tunnalH.add(cellP85.getNumericCellValue());
                        tunnalH.add(cellS85.getNumericCellValue());
                        double maxtunnalH = Collections.max(tunnalH);
                        double mintunnalH = Collections.min(tunnalH);
                        maxtunnalH = cellV85.getNumericCellValue();
                        mintunnalH = cellZ85.getNumericCellValue();
                        /*-----------------TunnalH-----------------*/
                        /*-----------------Amax-----------------*/
                        final ArrayList<Double> Amax = new ArrayList<>();
                        Amax.add(maxTunnalA);
                        Amax.add(maxtunnalB);
                        Amax.add(maxtunnalC);
                        Amax.add(maxtunnalD);
                        Amax.add(maxtunnalE);
                        Amax.add(maxtunnalF);
                        Amax.add(maxtunnalG);
                        Amax.add(maxtunnalH);
                        double resultAmax = Collections.max(Amax);
                        /*-----------------Amax-----------------*/
                        /*-----------------Amin-----------------*/
                        final ArrayList<Double> Amin = new ArrayList<>();
                        Amin.add(minTunnalA);
                        Amin.add(mintunnalB);
                        Amin.add(mintunnalC);
                        Amin.add(mintunnalD);
                        Amin.add(mintunnalE);
                        Amin.add(mintunnalF);
                        Amin.add(mintunnalG);
                        Amin.add(mintunnalH);
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
                textFieldSensitive.setText("");
                textFieldChannelDifference.setText("");
                labelTunnalMinResult.setText("");
                labelTunnalMaxResult.setText("");
                labelValueStabilityResult.setText("");
                labelValueToleranceResult.setText("");
                labelAbsorbanceRepeatResult.setText("");
                labelSensitiveResult.setText("");
                labelChannelDifferenceResult.setText("");
            }
            /*-----实验方法-----*/
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
                //count=dirs.length
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
//                System.out.println("关闭USB生产者监听...");
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
