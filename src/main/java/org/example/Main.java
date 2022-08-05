package org.example;

import com.spire.xls.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellRenderer;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.*;
import java.util.List;

public class Main {
    static JButton but1;
    static JButton but2;
    static JFileChooser fChooser;
    static JButton but3;
    static JButton but4;
    public static JFrame frame;
    static String multiFilePath = null;
    private static JTable table;
    private static CustomTableModel model;
    private int REFRESH = 0;// 刷新表格面板
    private static final String EXCEL_EXTENSION_XLSX = ".xlsx";// xlsx后缀
    private static final String EXCEL_EXTENSION_XLS = ".xls";// xls后缀
    private static final int MAXVALUE = 100;// 进度条最大值
    private static final int MINVALUE = 0;// 进度条最小值
    static int j = 0;// 判断是否点击保存按钮
    private static int MULTIPLE = 0;// 判断是否为多个文件

    /**
     * 类：列表中每行的属性
     */
    public static class FileRowModel{
        public File file;
        public int progress = 0;// 初始化进度条
        public int begin = 0;// 0 准备 1 完成
        public FileRowModel(File file) {
            this.file = file;
        }
        public int begin(){
            return begin;
        }
        public void setBegin(int value){
            begin = value;
        }
    }

    /**
     * 自定义列表模版
     */
    public static class CustomTableModel extends AbstractTableModel {

        public List<FileRowModel> rows = new ArrayList<>();

        // 初始化清空
        public void clear(){
            for (FileRowModel row : rows) {
                row.progress = 0;
                row.setBegin(0);
            }
            this.rows.clear();
        }
        // 添加行
        public void addRow(File file){
            rows.add(new FileRowModel(file));
            fireTableDataChanged();
        }
        // 更新每行进度条
        public void updateRow(FileRowModel fileRowModel){
            int i = rows.indexOf(fileRowModel);
            if(i >= 0){
                fireTableCellUpdated(i, 2);
            }
        }
        @Override
        public int getRowCount() {
            return rows.size();
        }
        @Override
        public int getColumnCount() {
            return 3;
        }
        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            if(columnIndex == 0){
                return rows.get(rowIndex).file.getName();
            }else if(columnIndex == 1){
                return rows.get(rowIndex).file.getAbsolutePath();
            }else{
                if(rows.get(rowIndex).begin() == 0){
                    return rows.get(rowIndex).progress;
                } else if (rows.get(rowIndex).begin() == 1) {
                    System.out.print("begin=1,filename="+rows.get(rowIndex).file.getName());
                    rows.get(rowIndex).progress = MAXVALUE;
                    return rows.get(rowIndex).progress;
                } else{
                    return 0;
                }
            }
        }
        // 设置表头
        public String getColumnName(int c) {
            if(c == 0)
                return "文件名";
            else if(c == 1)
                return "路径";
            else
                return "进度条";
        }
    }

    /**
     * 渲染进度条
     */
    public static class ProgressCellRender extends JProgressBar implements TableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            int progress = MINVALUE;
            if (value instanceof Float) {
                progress = Math.round(((Float) value) * 100f);
            } else if (value instanceof Integer) {
                progress = (int) value;
            }
            setValue(progress);
            return this;
        }
    }

    /**
     * 选择文件目录
     * @param index 索引
     * @param fileName 文件名
     * @return 返回导出路径
     */
    public String selectDir(int index, String fileName) {
        if (MULTIPLE == 0) {
            String oldFileName = model.rows.get(index).file.getName();
            String fromPath = model.rows.get(index).file.getPath().replace(oldFileName,fileName);// 选择当前文件夹
            fChooser = new JFileChooser();
            fChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            fChooser.setCurrentDirectory(new File(fromPath.replace(fileName,"")));
            fChooser.setDialogTitle("另存为");
            System.out.println(model.rows.size());
            // 仅有一个文件，显示保存文件菜单
            if (model.rows.size() == 1) {
                File f = new File(fileName);
                fChooser.setSelectedFile(f);
                int m = fChooser.showSaveDialog(frame);
                if (m == JFileChooser.APPROVE_OPTION) {
                    MULTIPLE = -1;
                    return fChooser.getSelectedFile().getAbsolutePath().replace("/"+fileName,"");
                }
            } else {
                j = fChooser.showOpenDialog(frame);
                if(j == JFileChooser.APPROVE_OPTION) {
                    File Path = fChooser.getSelectedFile();
                    fChooser.setSelectedFile(new File(fileName));
                    String toPath = Path.getPath().replace("/.","");
                    System.out.println(toPath);
                    MULTIPLE = 1;
                    multiFilePath = toPath;
                    return toPath;
                }
                MULTIPLE = -1;// 取消
            }
        } else {
            fChooser.setSelectedFile(new File(fileName));
            if(j == JFileChooser.APPROVE_OPTION){
                return multiFilePath;
            }
        }
        return "";
    }

    /**
     * 转换文件
     * @param index 索引
     * @param workbook 工作台
     * @param fileName 文件名称
     * @param fixed 后缀名
     */
    public static void processing(int index, Workbook workbook, String fileName, String fixed, String dest){
        new Thread(new Runnable() {
            @Override
            public void run() {
                Worksheet sheet = workbook.getWorksheets().get(0);
                String extension;
                // 访问工作表中使用的范围
                CellRange usedRange = sheet.getAllocatedRange();
                // 当将范围内的数字保存为文本时，忽略错误
                usedRange.setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));
                // 自适应行高、列宽
                usedRange.autoFitColumns();
                usedRange.autoFitRows();
                workbook.loadFromFile(model.rows.get(index).file.getAbsolutePath(),",",1,1);
                // 判断版本
                if(fixed.equals(EXCEL_EXTENSION_XLSX)){
                    workbook.saveToFile(fileName + fixed, ExcelVersion.Version2013);
                    extension = EXCEL_EXTENSION_XLSX;
                }else{
                    workbook.saveToFile(fileName + fixed, ExcelVersion.Version97to2003);
                    extension = EXCEL_EXTENSION_XLS;
                }
                // 移动文件
                String fromPath = System.getProperty("user.dir");
                fromPath = fromPath + '/' + fileName+fixed;
                File file = new File(fromPath);

                File d = new File(dest + '/' + fileName + extension);
                if(file.renameTo(d)){
                    System.out.println("Translate Success");
                }else{
                    System.out.println("False");
                }

                System.out.println("current complete, i="+index+",time="+System.currentTimeMillis());
                model.rows.get(index).setBegin(1);
            }
        }).start();
    }

    /**
     * 重复文件处理
     * @param path 路径
     * @return 返回布尔值
     */
    private static boolean duplicate(String path){
        for(int i = 0; i < model.rows.size();i++){
            if(model.rows.get(i).file.getAbsolutePath().equals(path)){
                JOptionPane.showMessageDialog(frame, "文件导入重复");
                return false;
            }
        }
        return true;
    }

    /**
     * 进度条更新
     * @param fileRowModel 行列表
     */
    private static void doProgressWork(FileRowModel fileRowModel){
        new Thread(new Runnable() {
            @Override
            public void run() {
                int i = 0;
                while(i < 97 && fileRowModel.begin == 0){
                    try {
                        Thread.sleep(new Random().nextInt(1000));
                    } catch (InterruptedException e) {
                        throw new RuntimeException(e);
                    }
                    i++;
                    fileRowModel.progress = i;
                    System.out.println("filename="+fileRowModel.file.getName()+",i="+i);
                    model.updateRow(fileRowModel);
                }
            }
        }).start();
    }

    /**
     * 更新GUI文件列表
     */
    private void refreshTable(){
        if(REFRESH != 0){
            model.clear();
            REFRESH = 0;
        }
    }

    /**
     * 创建并显示GUI。出于线程安全的考虑，
     * 这个方法在事件调用线程中调用。
     */
    public void createAndShowGUI() {

        // 创建及设置窗口
        frame = new JFrame("文件转换器");
        frame.setBounds(700,250,400,600);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        GridLayout gl = new GridLayout(5,5,5,3);
        frame.setLayout(gl);

        // 设置标题
        Panel pn0 = new Panel(new BorderLayout());
        java.net.URL imageURL1 = Main.class.getResource("/image/CSV.png");
        JLabel titleLabel = new JLabel(new ImageIcon(imageURL1));
        titleLabel.setBounds(300,100,10,100);
        pn0.add(titleLabel);
        frame.add(pn0);

        // 单选框组件
        Panel pn1 = new Panel();
        CheckboxGroup cg = new CheckboxGroup();
        Checkbox cg1 = new Checkbox("xlsx",cg,true);
        Checkbox cg2 = new Checkbox("xls",cg,false);
        pn1.add(cg1);
        pn1.add(cg2);
        frame.add(pn1);

        Container pane = frame.getContentPane();
        model = new CustomTableModel();
        table = new JTable(model);

        table.getColumnModel().getColumn(2).setCellRenderer(new ProgressCellRender());
        table.setSize(400, 300);
        table.setFillsViewportHeight(true);
        table.getColumnModel().getColumn(0).setMaxWidth(80);
        table.getColumnModel().getColumn(
                0).setMinWidth(60);
        table.getColumnModel().getColumn(1).setMinWidth(80);
        table.getColumnModel().getColumn(2).setMinWidth(80);
        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setBounds(0,0,200,400);
        pane.add(scrollPane);

        // 按钮1
        Panel pn2 = new Panel();
        but1 = new JButton("选择csv文件");
        but1.setContentAreaFilled(false);
        but1.setFont(new Font("隶书",Font.PLAIN,15));
        pn2.setLayout(new FlowLayout());
        pn2.add(but1);
        frame.add(pn2);

        // 按钮234
        Panel pn3 = new Panel();
        but2 = new JButton("转换");
        but3 = new JButton("清空");
        but4 = new JButton("移除");
        but2.setFont(new Font("隶书",Font.PLAIN,15));
        but3.setFont(new Font("隶书",Font.PLAIN,15));
        but4.setFont(new Font("隶书",Font.PLAIN,15));
        pn3.setLayout(new FlowLayout());
        pn3.add(but2);
        pn3.add(but3);
        pn3.add(but4);
        frame.add(pn3);

        // 拖动文件到文件框
        frame.setDropTarget(new DropTarget(){
            @Override
            public synchronized void drop(DropTargetDropEvent evt) {
                try{
                    refreshTable();// 初始化表格
                    evt.acceptDrop(DnDConstants.ACTION_COPY);
                    List<File> droppedFiles = (List<File>)evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    for (File file : droppedFiles) {
                        if(duplicate(file.getAbsolutePath())){
//                            System.out.println(file.getAbsolutePath());
                            if(file.getAbsolutePath().endsWith(".csv")){
                                model.addRow(file);
                            }else{
                                JOptionPane.showMessageDialog(frame,"请导入CSV文件");
                                throw new Exception();
                            }
                        }else{
                            throw new Exception();
                        }
                    }
                }catch (Exception ignored){

                }
            }
        });
        frame.setVisible(true);

        // 监听选择文件按钮
        but1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                chooser.setFileFilter(new FileNameExtensionFilter("csv(*.csv)", "csv"));
                chooser.setMultiSelectionEnabled(true);
                refreshTable();// 初始化表格
                int ret = chooser.showDialog(frame, "选择");
                if (ret == JFileChooser.APPROVE_OPTION){
                    File[] files = chooser.getSelectedFiles();
                    if(files.length == 0){
                        return;
                    }
                    try{
                        for (File file : files) {
                            if (duplicate(file.getPath())) {
                                if (file.getAbsolutePath().endsWith(".csv")) {
                                    model.addRow(file);
                                }
                            }
                        }
                    }catch (Exception e2){
                        JOptionPane.showMessageDialog(frame,"导入文件重复");
                    }
                }
            }
        });

        // 选择文件按钮
        but2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try{
                    MULTIPLE = 0;
                    // 判断文件是否为空
                    if(model.rows.isEmpty()){
                        throw new RuntimeException();
                    }
                    for(int i = 0; i < model.getRowCount(); i++){
                        // 创建一个workbook
                        Workbook workbook = new Workbook();
                        // 从列表中依次提取单个文件名
                        String fileName = model.rows.get(i).file.getName().substring(0, model.rows.get(i).file.getName().lastIndexOf("."));
                        // 版本： 1表示xlsx 2表示xls
                        if(cg1.getState()){
                            String dest = selectDir(i, fileName+EXCEL_EXTENSION_XLSX);
                            if(dest.length() != 0){
                                processing(i, workbook, fileName, EXCEL_EXTENSION_XLSX, dest);
                                doProgressWork(model.rows.get(i));
                                REFRESH = 1;// 再度拖入文件后刷新文件表格
                            }
                        }else{
                            String dest = selectDir(i, fileName+EXCEL_EXTENSION_XLSX);
                            if(dest.length() != 0){
                                processing(i, workbook, fileName, EXCEL_EXTENSION_XLS, dest);
                                doProgressWork(model.rows.get(i));
                                REFRESH = 1;// 再度拖入文件后刷新文件表格
                            }
                        }
                    }
                } catch (Exception e1){
                    e1.printStackTrace();
                    JOptionPane.showMessageDialog(frame,"未导入文件");
                }
            }
        });

        // 监听清空按钮
        but3.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                model.clear();
                REFRESH = 0;
                table.updateUI();
                JOptionPane.showMessageDialog(frame,"已清空");
            }
        });

        // 监听移除按钮
        but4.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try{
                    int[] rows = table.getSelectedRows();
                    // 判断是否选中为空
                    int l = rows.length;
                    for(int i = l-1; i >= 0; i--) {
                        model.rows.remove(rows[i]);
                        table.updateUI();
                    }
                }catch (Exception e1){
                    JOptionPane.showMessageDialog(frame,"未选中");
                }
            }
        });
    }

    public Main(){
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }

    public static void main(String[] args) {
        new Main();
    }
}
