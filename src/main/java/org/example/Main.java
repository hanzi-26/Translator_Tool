package org.example;
import com.spire.xls.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.util.*;
import java.awt.*;
import javax.swing.*;
import java.io.File;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JProgressBar;
import java.awt.Color;
import java.util.List;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;

import static java.lang.Thread.sleep;

public class Main{
    // 窗口所需组件
    private static JFrame f;
    private File cfile;
    private volatile JProgressBar progressBar;
    private static final int PROGRESS_MIN_VALUE = 0;
    private static final int PROGRESS_MAX_VALUE = 100;
    JButton but1;
    JButton but2;
    JButton but3;
    JButton but4;

    JTextField p1;
    Checkbox cg1;
    Checkbox cg2;
    JLabel titleLabel;
    JFileChooser fChooser;
    private static final String EXTENSION_XLSX = ".xlsx";
    private static final String EXTENSION_XLS = ".xls";
    private int Begin = 0;// 0 准备 1 完成
    // 文件表格所需参数
    private String filepath;// 文件路径
    List<String> fileListPath = new ArrayList<>();// 文件储存路径列表
    List<String> fileListName = new ArrayList<>();// 文件名储存列表
    private static int REFRESH = 0;// 判断是否转换过数据
    private static int MULTIPLE = 0;// 判断是否为多个文件
    private  String multiFilePath;// 目标文件路径
    File Path;
    int j;
    JTable table;

    /**
     *  绘制进度条到表格中
     */
    class ProgressCellRender extends JProgressBar implements TableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
//            progressBar.paintImmediately(progressBar.getBounds());
            int progress = progressBar.getValue();
            if(value instanceof Float){
                progress = Math.round(((Float) value)*100f);
            } else if(value instanceof Integer){
                progress = (int)value;
            }
            setValue(progress);
            return this;
        }
    }

    //    class ProgressRenderer extends DefaultTableCellRenderer {
//
//        private final JProgressBar b = new JProgressBar(0, 100);
//
//        public ProgressRenderer() {
//            super();
//            setOpaque(true);
//            b.setBorder(BorderFactory.createEmptyBorder(1, 1, 1, 1));
//        }
//
//        @Override
//        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
//            Integer i = (Integer) value;
//            String text = "Completed";
//            if (i < 0) {
//                text = "Error";
//            } else if (i < 100) {
//                b.setValue(i);
//                return b;
//            }
//            super.getTableCellRendererComponent(table, text, isSelected, hasFocus, row, column);
//            return this;
//        }
//    }
    String[][] datas = {};
    String[] titles = {"文件名", "路径", "进度条"};
    private DefaultTableModel model = new DefaultTableModel(datas, titles) {
        private static final long serialVersionUID = 1L;

        @Override
        public Class<?> getColumnClass(int column) {
            return getValueAt(0, column).getClass();
        }

        @Override
        public boolean isCellEditable(int row, int col) {
            return false;
        }
    };

    /**
     * GUI界面
     */
    public void frameWindow(){
        f = new JFrame("表格转换工具");
        f.setBounds(500, 300, 400, 600);

        // 设置窗口最小纬度
        Dimension dim = new Dimension(200,500);
        f.setMinimumSize(dim);
        // 创建菜单容器
        MenuBar mb = new MenuBar();
        f.setMenuBar(mb);

        // 设置表格4行2列的面板
        GridLayout gl = new GridLayout(5,5,5,3);
        f.setLayout(gl);

        // 设置标题
        Panel pn0 = new Panel(new BorderLayout());
        titleLabel = new JLabel(new ImageIcon("CSV.png"));
        titleLabel.setBounds(300,100,10,100);
        pn0.add(titleLabel);
//        pn0.setBackground(new Color(67, 67, 68));
        pn0.setLayout(new FlowLayout());
        f.add(pn0);

        // 单选框组件
        Panel pn1 = new Panel();
        CheckboxGroup cg = new CheckboxGroup();
        pn1.setLayout(new FlowLayout());
        cg1 = new Checkbox("xlsx",cg,true);
        cg2 = new Checkbox("xls",cg,false);
//        pn1.setBackground(new Color(67, 67, 68));
        pn1.add(cg1);
        pn1.add(cg2);
        f.add(pn1);

        // 创建表哥所在面板
        Panel pn5 = new Panel(new BorderLayout());
        p1 = new JTextField(0);
        p1.setBounds(20,20,3,100);

        model = new DefaultTableModel(datas, titles);

        table = new JTable(model);
        table.getColumn("进度条").setCellRenderer(new ProgressCellRender());
        table.getColumnModel().getColumn(0).setMaxWidth(80);
        table.getColumnModel().getColumn(
                0).setMinWidth(60);
        table.getColumnModel().getColumn(1).setMinWidth(80);
        table.getColumnModel().getColumn(2).setMinWidth(80);
//        pn5.setBackground(new Color(67, 67, 68));

        table.paintImmediately(table.getBounds());
        JScrollPane scroll = new JScrollPane(table,JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scroll.setBounds(0,0,200,400);
        scroll.getViewport().setBackground(new Color(243,247,255));

        pn5.add(p1);
        pn5.add(table);
        pn5.add(new JScrollPane(table));
        table.setSize(400, 300);
        f.add(pn5);

        // 按钮1
        Panel pn2 = new Panel();
        but1 = new JButton("选择csv文件");
        but1.setContentAreaFilled(false);
        but1.setFont(new Font("隶书",Font.PLAIN,15));
        pn2.setLayout(new FlowLayout());
        pn2.add(but1);
        f.add(pn2);

        // 按钮2
        Panel pn3 = new Panel();
        but2 = new JButton("转换");
        but3 = new JButton("清空");
        but4 = new JButton("移除");
        but2.setFont(new Font("隶书",Font.PLAIN,15));
        but3.setFont(new Font("隶书",Font.PLAIN,15));
        but4.setFont(new Font("隶书",Font.PLAIN,15));
//        pn3.setBackground(new Color(67, 67, 68));
        pn3.setLayout(new FlowLayout());
        pn3.add(but2);
        pn3.add(but3);
        pn3.add(but4);
        f.add(pn3);

        // 进度条
        progressBar = new JProgressBar();
        progressBar.setMaximum(PROGRESS_MAX_VALUE);
        progressBar.setMinimum(PROGRESS_MIN_VALUE);
        progressBar.setForeground(new Color(46, 145, 228));
        progressBar.setBackground(new Color(220, 220, 220));

        // 拖动窗口
        f.setDropTarget(new DropTarget(){
            @Override
            public synchronized void drop(DropTargetDropEvent evt) {
                try{
                    // 转换完成后再次添加文件时，更新文件列表
                    refreshTable();
                    evt.acceptDrop(DnDConstants.ACTION_COPY);
                    List<File> droppedFiles = (List<File>)evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    for (File files : droppedFiles) {
                        filepath = files.getAbsolutePath();
                    }
                    if(filepath.endsWith(".csv")){
                        duplicate(fileListPath,filepath);
                        fileListPath.add(filepath);
                        cfile = new File(filepath);
                        fileListName.add(cfile.getName());
                        titleLabel.setIcon(new ImageIcon("excel.png"));
                        model.addRow(new Object[] { cfile.getName(), cfile.getAbsolutePath() });
                    } else {
                        initGUI("请导入CSV文件！");
                    }
                }catch (Exception e){
                    initGUI("文件导入重复！");
                }
            }
        });
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        f.setVisible(true);
    }

    /**
     * 避免重复
     * @param lists
     * @param Path
     */
    private static void duplicate(List<String> lists, String Path){
        if(lists.contains(Path)){
            lists.remove(lists.size());
        }
    }

    /**
     * 重置常数
     */
    private void initConstant(){
        MULTIPLE = 0;
        Begin = 0;
        progressBar.setValue(PROGRESS_MIN_VALUE);
    }

    /**
     * 更新GUI文件列表
     */
    private void refreshTable(){
        if(REFRESH != 0){
            progressBar.setValue(PROGRESS_MIN_VALUE);
            model.getDataVector().clear();
            fileListPath.clear();
            fileListName.clear();
            REFRESH = 0;
        }
    }

    public void moveFile(String fileName){
        String fromPath = System.getProperty("user.dir");
        System.out.println(fileName);
        fromPath = fromPath + '/' + fileName;
        File file = new File(fromPath);
        fChooser.setSelectedFile(new File(fileName));
        if(j == JFileChooser.APPROVE_OPTION){
            String toPath = multiFilePath;
            toPath = toPath + fileName;
            File toFile = new File(toPath);
            System.out.println(fromPath+"\t"+toPath);
            System.out.println(file.renameTo(toFile));
            file.renameTo(toFile);
        }
    }

    /**
     * 将文件移动到指定文件夹下
     * @param fileName
     * @param customCallback
     * @throws InterruptedException
     */
    public void moveFile(String fileName, CustomCallback customCallback) throws InterruptedException {
        //  导入到同一个文件目录下
        // 转移文件所需参数
        // 目标文件
        File toFile;
        if(MULTIPLE == 0){
            String fromPath = System.getProperty("user.dir");
            fromPath = fromPath + '/' + fileName;
            File file = new File(fromPath);
            // 用户调取地址
            fChooser = new JFileChooser();
            String chooserPath = cfile.getPath();
            chooserPath = chooserPath.replace(cfile.getName(),"");
            fChooser.setCurrentDirectory(new File(chooserPath));// 设置默认目录
            fChooser.setDialogTitle("另存为"); // 定义目录框标题
            fChooser.setSelectedFile(new File(fileName));
            // 是否点击保存按钮
            j = fChooser.showSaveDialog(f);
            if(j == JFileChooser.APPROVE_OPTION) {
                Path = fChooser.getSelectedFile();
                String toPath = Path.getPath();
                multiFilePath = toPath.replace(fileName, "");
                toFile = new File(toPath);
                Boolean b = file.renameTo(toFile);
                System.out.println(b);
                file.renameTo(toFile);
                Thread t;
                t = new Thread(new Runnable() {
                    @Override
                    public void run() {
                        try {
                            while(Begin != 0){
                                sleep(1000);
                                System.out.println("Waiting");
                            }
                        } catch (InterruptedException e) {
                            throw new RuntimeException(e);
                        }
                    }
                });
                t.start();
                System.out.println("Running....");
                customCallback.ok();
            }
            MULTIPLE = 1;
        } else {
            String fromPath = System.getProperty("user.dir");
            System.out.println(fileName);
            fromPath = fromPath + '/' + fileName;
            File file = new File(fromPath);
            fChooser.setSelectedFile(new File(fileName));
            if(j == JFileChooser.APPROVE_OPTION){
                String toPath = multiFilePath;
                toPath = toPath + fileName;
                toFile = new File(toPath);
                System.out.println(fromPath+"\t"+toPath);
                System.out.println(file.renameTo(toFile));
                file.renameTo(toFile);
                customCallback.ok();
            }
        }
    }

    /**
     * 初始化GUI窗口
     * @param descrip
     */
    public void initGUI(String descrip){
        titleLabel.setIcon(new ImageIcon("CSV.png"));
        progressBar.setValue(PROGRESS_MIN_VALUE);
        JOptionPane.showMessageDialog(f, descrip);
    }

    /**
     * 转换文件
     * @param workbook
     * @param fixed
     */
    public void processing(Workbook workbook, String fileName, String fixed){
        new Thread(new Runnable() {
            @Override
            public void run() {
                Worksheet sheet = workbook.getWorksheets().get(0);
                // 访问工作表中使用的范围
                CellRange usedRange = sheet.getAllocatedRange();
                // 当将范围内的数字保存为文本时，忽略错误
                usedRange.setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));
                // 自适应行高、列宽
                usedRange.autoFitColumns();
                usedRange.autoFitRows();
                workbook.loadFromFile(cfile.getPath(),",",1,1);
                // 判断版本
                if(fixed.equals(EXTENSION_XLSX)){
                    workbook.saveToFile(fileName + fixed, ExcelVersion.Version2013);
                }else{
                    workbook.saveToFile(fileName + fixed, ExcelVersion.Version97to2003);
                }
                Begin = 1;
                setValues(PROGRESS_MAX_VALUE);
                table.updateUI();
            }
        }).start();
    }

    /**
     * 避免进度条线程冲突
     * @param val
     */
    private synchronized void setValues(int val){
        if(progressBar.getValue() != PROGRESS_MAX_VALUE)
            progressBar.setValue(val);
    }

    /**
     * 进度条
     * @return
     */
    public void processBar(){
        new Thread(new Runnable() {
            @Override
            public void run() {
                int i = PROGRESS_MIN_VALUE;
                while(i < 97){
                    if(Begin == 1) {
                        setValues(PROGRESS_MAX_VALUE);
                        break;
                    }else{
                        try {
                            setValues(i);
                            if(i < 40){
                                sleep(100);
                                table.updateUI();
                            }else if(i < 60){
                                sleep(150);
                                table.updateUI();
                            }else if(i < 80){
                                sleep(250);
                                table.updateUI();
                            }else {
                                sleep(550);
                                table.updateUI();
                            }
                        } catch (InterruptedException ignored) {

                        }
                    }
                    i++;
                }
            }
        }).start();
    }

    public Main(){

        // GUI界面
        frameWindow();

        // 监听选择文件按钮
        but1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                chooser.setFileFilter(new FileNameExtensionFilter("csv(*.csv)", "csv"));
                chooser.setMultiSelectionEnabled(true);
                refreshTable();
                int ret = chooser.showDialog(f, "选择");
                if (ret == JFileChooser.APPROVE_OPTION){
                    File[] files = chooser.getSelectedFiles();
                    if(files.length == 0){
                        return;
                    }
                    try{
                        for(int i = 0; i < files.length; i++){
                            cfile = files[i];
                            duplicate(fileListPath,cfile.getPath());
                            fileListPath.add(cfile.getPath());
                            fileListName.add(cfile.getName());
                            model.addRow(new Object[] { cfile.getName(), cfile.getAbsolutePath() });
                        }
                    }catch (Exception e2){
                        initGUI("导入文件重复！");
                    }
                }
            }
        });

        // 监听转换按钮
        but2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try{
                    initConstant();
                    // 判断文件是否为空
                    if(fileListPath.isEmpty()){
                        throw new RuntimeException();
                    }
                    for(int i = 0; i < fileListPath.size(); i++){
                        cfile = new File(fileListPath.get(i));
                        Begin = 0;
                        setValues(PROGRESS_MIN_VALUE);
                        // 创建一个workbook
                        Workbook workbook = new Workbook();
                        // 从列表中依次提取单个文件名
                        String fileName = fileListName.get(i).substring(0, cfile.getName().lastIndexOf("."));
                        // 版本： 1表示xlsx 2表示xls
                        if(cg1.getState()){
                            moveFile(fileName + ".xlsx", new CustomCallback() {
                                @Override
                                public void ok() throws InterruptedException {
                                    processing(workbook, fileName, EXTENSION_XLSX);
                                    processBar();
                                    setValues(PROGRESS_MIN_VALUE);
                                    REFRESH = 1;// 再度拖入文件后刷新文件表格
                                    sleep(100);// 等待计算完成
                                }
                            });
                            //TODO
                            // RenameTo速度可能比moveFile慢，导致先发生moveFile的情况
                            moveFile(fileName+".xlsx");
                        }else{
                            moveFile(fileName + ".xls", new CustomCallback() {
                                @Override
                                public void ok() throws InterruptedException {
                                    processing(workbook, fileName, EXTENSION_XLS);
                                    processBar();
                                    setValues(PROGRESS_MIN_VALUE);
                                    REFRESH = 1;// 再度拖入文件后刷新文件表格
                                    sleep(100);// 等待计算完成
                                }
                            });
                            //TODO
                            // RenameTo速度可能比moveFile慢，导致先发生moveFile的情况
                            moveFile(fileName+".xls");
                        }
                    }

                } catch (Exception e1){
                    initGUI("未导入文件！");
                }
            }
        });

        // 监听清空按钮
        but3.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cfile = null;
                model.getDataVector().clear();
                fileListPath.clear();
                fileListName.clear();
                initGUI("已清空");
            }
        });
        // 监听移除按钮
        but4.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try{
                    int rows[] = table.getSelectedRows();
                    int[] empty = {};
                    // 判断是否选中为空
                    if(Arrays.toString(rows).equals(Arrays.toString(empty))){
//                        System.out.println("ffff");
                        throw new RuntimeException();
                    }
                    for(int i = rows.length-1; i >= 0; i--) {
                        fileListPath.remove(rows[i]);
                        fileListName.remove(rows[i]);
                        model.removeRow(rows[i]);
//                        System.out.println(rows[i]);
                    }
//                    System.out.println(fileListPath);
//                    System.out.println(fileListName);
                }catch (Exception e1){
                    JOptionPane.showMessageDialog(f,"未选中");
                }
            }
        });

    }

    public static void main(String[] args) {
        // TODO Auto-generated method stub
        new Main();
    }

}


