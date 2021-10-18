import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Set;
import java.util.TreeMap;
import net.proteanit.sql.DbUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.math3.util.ArithmeticUtils;
public class Contact {
    private JPanel Main;
    private JTextField txtName;
    private JTextField txtMobile;
    private JButton saveButton;
    private JTable table1;
    private JButton updateButton;
    private JButton deleteButton;
    private JButton searchButton;
    private JTextField txtId;
    private JScrollPane table_1;
    private JButton button1;
    private JButton exporttoexcelButton;
    private JTextField txtAddress;
    private JTextField txtGender;
    private JTextField txtEmail;

    public static void main(String[] args) {
        JFrame frame = new JFrame("Contact");
        frame.setContentPane(new Contact().Main);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }
    Connection con;
    PreparedStatement pst;
    public void connect()
    {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            con = DriverManager.getConnection("jdbc:mysql://localhost/DBContact", "root", "");
            System.out.println("Success");
        } catch (ClassNotFoundException ex) {
            ex.printStackTrace();

        } catch (SQLException ex) {

            ex.printStackTrace();

        }
    }
    void table_load()
    {
        try {
            pst = con.prepareStatement("select * from phcontact");
            ResultSet rs = pst.executeQuery();
            table1.setModel(DbUtils.resultSetToTableModel(rs));
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
    public Contact() {
        connect();
        table_load();
        //save the details
        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String conname,mobile,address,gender,email;
                conname=txtName.getText();
                mobile=txtMobile.getText();
                address=txtAddress.getText();
                gender=txtGender.getText();
                email=txtEmail.getText();
                try{

                    pst=con.prepareStatement("insert into phcontact(contname,mobile,address,gender,email) values (?,?,?,?,?)");
                    pst.setString(1,conname);
                    pst.setString(2,mobile);
                    pst.setString(3,address);
                    pst.setString(4,gender);
                    pst.setString(5,email);

                    pst.executeUpdate();
                    JOptionPane.showMessageDialog(null,"Record Added!!!");
                    table_load();
                    txtName.setText("");
                    txtMobile.setText("");
                    txtAddress.setText("");
                    txtGender.setText("");
                    txtEmail.setText("");
                    txtName.requestFocus();
                }
                catch (SQLException e1){
                    e1.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Name already exists in the database!!! Please enter a new Name");
                }
            }
        });
        //search the details by name
        searchButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    String searchid = txtId.getText();
                    pst = con.prepareStatement("SELECT contname,mobile,address,gender,email FROM `phcontact` WHERE contname = ?");
                    pst.setString(1, searchid);
                    ResultSet rs = pst.executeQuery();
                    if (rs.next()==true) {
                        String cname= rs.getString(1);
                        String cmobile = rs.getString(2);
                        String caddress = rs.getString(3);
                        String cgender = rs.getString(4);
                        String cemail = rs.getString(5);
                        txtName.setText(cname);
                        txtMobile.setText(cmobile);
                        txtAddress.setText(caddress);
                        txtGender.setText(cgender);
                        txtEmail.setText(cemail);

                    }
                    else
                    {
                        txtName.setText("");
                        txtMobile.setText("");
                        txtAddress.setText("");
                        txtGender.setText("");
                        txtEmail.setText("");
                        txtName.requestFocus();
                        JOptionPane.showMessageDialog(null, "Invalid name");
                    }
                }
                catch (SQLException ex)
                {
                    ex.printStackTrace();
                }
            }
        });
        //update the changes of the searched contacts
        updateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String name,mobile,add,gen,em,eid;
                name=txtName.getText();
                mobile=txtMobile.getText();
                add=txtAddress.getText();
                gen=txtGender.getText();
                em=txtEmail.getText();
                eid=txtId.getText();
                try{
                    pst=con.prepareStatement("UPDATE phcontact SET contname=?,mobile=?,address=?,gender=?,email=? WHERE contname = ?");
                    pst.setString(1,name);
                    pst.setString(2,mobile);
                    pst.setString(3,add);
                    pst.setString(4,gen);
                    pst.setString(5,em);
                    pst.setString(6,eid);
                    pst.executeUpdate();
                    JOptionPane.showMessageDialog(null,"Updated!!!");
                    table_load();
                    txtName.setText("");
                    txtMobile.setText("");
                    txtAddress.setText("");
                    txtGender.setText("");
                    txtEmail.setText("");
                    txtName.requestFocus();
                    txtName.requestFocus();
                }
                catch (SQLException e1)
                {
                    e1.printStackTrace();
                }
            }
        });
        //delete the details of the searched contacts
        deleteButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String eid;
                eid=txtId.getText();
                try{
                    pst=con.prepareStatement("delete from phcontact where contname =?");
                    pst.setString(1,eid);
                    pst.executeUpdate();
                    JOptionPane.showMessageDialog(null,"Deleted");
                    table_load();
                    txtName.setText("");
                    txtMobile.setText("");
                    txtAddress.setText("");
                    txtGender.setText("");
                    txtEmail.setText("");
                    txtName.requestFocus();
                    txtName.requestFocus();
                }
                catch (SQLException e1) {
                    e1.printStackTrace();
                }
            }
        });
        //create and save in excel file
        exporttoexcelButton.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    try{

                        Statement statement = con.createStatement();
                        FileOutputStream fileOut;
                        fileOut = new FileOutputStream("file.xls");
                        HSSFWorkbook workbook = new HSSFWorkbook();
                        HSSFSheet worksheet = workbook.createSheet("Sheet 0");
                        Row row1 = worksheet.createRow((short)0);
                        row1.createCell(0).setCellValue("id");
                        row1.createCell(1).setCellValue("contname");
                        row1.createCell(2).setCellValue("mobile");
                        row1.createCell(3).setCellValue("address");
                        row1.createCell(4).setCellValue("gender");
                        row1.createCell(5).setCellValue("email");
                        Row row2 ;
                        ResultSet rs = statement.executeQuery("SELECT id, contname, mobile,address,gender,email FROM phcontact");
                        while(rs.next()){
                            int a = rs.getRow();
                            row2 = worksheet.createRow((short)a);
                            row2.createCell(0).setCellValue(rs.getString(1));
                            row2.createCell(1).setCellValue(rs.getString(2));
                            row2.createCell(2).setCellValue(rs.getString(3));
                            row2.createCell(3).setCellValue(rs.getString(4));
                            row2.createCell(4).setCellValue(rs.getString(5));
                            row2.createCell(5).setCellValue(rs.getString(6));
                        }
                        workbook.write(fileOut);
                        fileOut.flush();
                        fileOut.close();
                        rs.close();
                        statement.close();
                        System.out.println("Export Success");
                    }
                    catch(SQLException ex){
                        System.out.println(ex);
                    }catch(IOException ioe){
                        System.out.println(ioe);
                    }
                }
            });
        }
    }
