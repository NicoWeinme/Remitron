import java.awt.Desktop;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.time.*;

public class Ventana extends JFrame {
    private fondoPanel panel = new fondoPanel();
    private JButton botonPrint = new JButton();
    private JButton botonExit = new JButton();
    private JTextField DIA = new JTextField();
    private JTextField MES = new JTextField();
    private JTextField AÑO = new JTextField();
    private JTextField NOMBRE = new JTextField();
    private JTextField CUIT = new JTextField();
    private JTextField DOMICILIO = new JTextField();
    private JTextField CONDICION = new JTextField();
    private JTextField RAZON = new JTextField();
    private JTextField NOMBRET = new JTextField();
    private JTextField CUITT = new JTextField();
    private JTextField DNI = new JTextField();
    private JTextField CANT= new JTextField();
    private JTextField UN = new JTextField();
    private JTextField DESC = new JTextField();
    private JTextField CANT1= new JTextField();
    private JTextField UN1 = new JTextField();
    private JTextField DESC1 = new JTextField();
    private JTextField CANT2= new JTextField();
    private JTextField UN2 = new JTextField();
    private JTextField DESC2 = new JTextField();
    private JTextField CANT3= new JTextField();
    private JTextField UN3 = new JTextField();
    private JTextField DESC3 = new JTextField();
    private JTextField CANT4= new JTextField();
    private JTextField UN4 = new JTextField();
    private JTextField DESC4 = new JTextField();
    private JTextField CANT5= new JTextField();
    private JTextField UN5 = new JTextField();
    private JTextField DESC5 = new JTextField();
    private JTextField CANT6= new JTextField();
    private JTextField UN6 = new JTextField();
    private JTextField DESC6 = new JTextField();
    private JTextField CANT7= new JTextField();
    private JTextField UN7 = new JTextField();
    private JTextField DESC7 = new JTextField();
    private JTextField CANT8= new JTextField();
    private JTextField UN8 = new JTextField();
    private JTextField DESC8 = new JTextField();
    private JTextField CANT9= new JTextField();
    private JTextField UN9 = new JTextField();
    private JTextField DESC9 = new JTextField();

    private LocalDate fechaActual = LocalDate.now();

    public Ventana(){
        this.setSize(500,750);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setTitle("Remitron");
        setLocationRelativeTo(null);
        iniciando();


    }
    private void iniciando(){
        Panel();
        Etiquetas();
        CamposTexto();
        Botones();
    }

    private void Panel(){
        panel = new fondoPanel();
        panel.setLayout(null);
        this.getContentPane().add(panel);
    }
    class fondoPanel extends JPanel{
        private Image imagen;
        @Override
        public void paint(Graphics g){
            imagen = new ImageIcon(getClass().getResource("/Images/FondoPanelPrincipal.jpg")).getImage();
            g.drawImage(imagen, 0,0, panel.getWidth(),panel.getHeight(),this);
            setOpaque(false);
            super.paint(g);
        }
    }
    private void Etiquetas(){

    }
    private void CamposTexto(){
        DIA();
        MES();
        AÑO();
        NOMBRE();
        CUIT();
        DOMICILIO();
        CONDICION();
        RAZON();
        CUITT();
        NOMBRET();
        DNI();
        CANT();
        UN();
        DESC();

    }
    private void DIA(){
        JLabel LDIA = new JLabel("DÍA", SwingConstants.RIGHT);
        LDIA.setBounds(35,10,30,25);
        panel.add(LDIA);
        DIA = new JTextField();
        DIA.setText(String.valueOf(fechaActual.getDayOfMonth()));
        DIA.setBounds(70,10,25,25);
        panel.add(DIA);
    }
    private void MES(){
        JLabel LMES = new JLabel("MES", SwingConstants.RIGHT);
        LMES.setBounds(110,10,30,25);
        panel.add(LMES);
        MES = new JTextField();
        MES.setText(String.valueOf(fechaActual.getMonthValue()));
        MES.setBounds(145,10,25,25);
        panel.add(MES);

    }
    private void AÑO(){
        JLabel LAÑO = new JLabel("AÑO", SwingConstants.RIGHT);
        LAÑO.setBounds(185,10,30,25);
        panel.add(LAÑO);
        AÑO = new JTextField();
        String ANIO;
        AÑO.setText(String.valueOf(fechaActual.getYear()));
        ANIO = AÑO.getText().charAt(2)+""+AÑO.getText().charAt(3);
        AÑO.setText(ANIO);
        AÑO.setBounds(220,10,25,25);
        panel.add(AÑO);
    }
    private void NOMBRE(){
        JLabel LNOMBRE = new JLabel("NOMBRE", SwingConstants.RIGHT);
        LNOMBRE.setBounds(15,50,50,25);
        panel.add(LNOMBRE);
        NOMBRE = new JTextField();
        NOMBRE.setText("Tito Ibañez");
        NOMBRE.setBounds(75,50,170,25);
        panel.add(NOMBRE);
    }
    private void CUIT(){
        JLabel LCUIT = new JLabel("CUIT", SwingConstants.RIGHT);
        LCUIT.setBounds(240,50,50,25);
        panel.add(LCUIT);
        CUIT = new JTextField();
        CUIT.setText("11-22222222-33");
        CUIT.setBounds(295,50,170,25);
        panel.add(CUIT);
    }
    private void DOMICILIO(){
        JLabel LDOMICILIO = new JLabel("DOMICILIO", SwingConstants.LEFT);
        LDOMICILIO.setBounds(13,90,70,25);
        panel.add(LDOMICILIO);
        DOMICILIO = new JTextField();
        DOMICILIO.setText("AV. Las Tipas y los Tipos 987");
        DOMICILIO.setBounds(75,90,170,25);
        panel.add(DOMICILIO);
    }
    private void CONDICION(){
        JLabel LCONDICION = new JLabel("CONDICIÓN IVA", SwingConstants.LEFT);
        LCONDICION.setBounds(13,130,90,25);
        panel.add(LCONDICION);
        CONDICION = new JTextField();
        CONDICION.setBounds(100,130,170,25);
        panel.add(CONDICION);
    }
    private void RAZON(){
        JLabel LRAZON = new JLabel("RAZÓN SOCIAL", SwingConstants.LEFT);
        LRAZON.setBounds(13,170,90,25);
        panel.add(LRAZON);
        RAZON = new JTextField();
        RAZON.setText("Ecogas SA");
        RAZON.setBounds(100,170,170,25);
        panel.add(RAZON);
    }
    private void NOMBRET(){
        JLabel LNOMBRET = new JLabel("CONDUCTOR", SwingConstants.LEFT);
        LNOMBRET.setBounds(13,210,90,25);
        panel.add(LNOMBRET);
        NOMBRET = new JTextField();
        NOMBRET.setText("El Pelado del transportador");
        NOMBRET.setBounds(100,210,170,25);
        panel.add(NOMBRET);
    }
    private void CUITT(){
        JLabel LCUITT = new JLabel("CUIT", SwingConstants.RIGHT);
        LCUITT.setBounds(15,250,50,25);
        panel.add(LCUITT);
        CUITT = new JTextField();
        CUITT.setBounds(70,250,170,25);
        panel.add(CUITT);
    }
    private void DNI(){
        JLabel LDNI = new JLabel("DNI", SwingConstants.RIGHT);
        LDNI.setBounds(235,250,50,25);
        panel.add(LDNI);
        DNI = new JTextField();
        DNI.setBounds(295,250,170,25);
        panel.add(DNI);
    }
    private void CANT(){
        CANT = new JTextField();
        JLabel LCANT = new JLabel("CANT", SwingConstants.CENTER);
        LCANT.setBounds(15,295,50,25);
        panel.add(LCANT);
        CANT.setBounds(15,320,50,25);
        CANT1 = new JTextField();
        CANT1.setBounds(15,355,50,25);
        CANT2 = new JTextField();
        CANT2.setBounds(15,390,50,25);
        CANT3 = new JTextField();
        CANT3.setBounds(15,425,50,25);
        CANT4 = new JTextField();
        CANT4.setBounds(15,460,50,25);
        CANT5 = new JTextField();
        CANT5.setBounds(15,495,50,25);
        CANT6 = new JTextField();
        CANT6.setBounds(15,530,50,25);
        CANT7 = new JTextField();
        CANT7.setBounds(15,565,50,25);
        CANT8 = new JTextField();
        CANT8.setBounds(15,600,50,25);
        CANT9 = new JTextField();
        CANT9.setBounds(15,635,50,25);
        panel.add(CANT);
        panel.add(CANT1);
        panel.add(CANT2);
        panel.add(CANT3);
        panel.add(CANT4);
        panel.add(CANT5);
        panel.add(CANT6);
        panel.add(CANT7);
        panel.add(CANT8);
        panel.add(CANT9);
    }
    private void UN(){
        UN = new JTextField();
        JLabel LUN = new JLabel("UN", SwingConstants.CENTER);
        LUN.setBounds(70,295,50,25);
        panel.add(LUN);
        UN.setBounds(70,320,50,25);
        UN1 = new JTextField();
        UN1.setBounds(70,355,50,25);
        UN2 = new JTextField();
        UN2.setBounds(70,390,50,25);
        UN3 = new JTextField();
        UN3.setBounds(70,425,50,25);
        UN4 = new JTextField();
        UN4.setBounds(70,460,50,25);
        UN5 = new JTextField();
        UN5.setBounds(70,495,50,25);
        UN6 = new JTextField();
        UN6.setBounds(70,530,50,25);
        UN7 = new JTextField();
        UN7.setBounds(70,565,50,25);
        UN8 = new JTextField();
        UN8.setBounds(70,600,50,25);
        UN9 = new JTextField();
        UN9.setBounds(70,635,50,25);
        panel.add(UN);
        panel.add(UN1);
        panel.add(UN2);
        panel.add(UN3);
        panel.add(UN4);
        panel.add(UN5);
        panel.add(UN6);
        panel.add(UN7);
        panel.add(UN8);
        panel.add(UN9);
    }
    private void DESC(){
        DESC = new JTextField();
        JLabel LDESC = new JLabel("DESCRIPCIÓN", SwingConstants.CENTER);
        LDESC.setBounds(125,295,300,25);
        panel.add(LDESC);
        DESC.setBounds(125,320,300,25);
        DESC1 = new JTextField();
        DESC1.setBounds(125,355,300,25);
        DESC2 = new JTextField();
        DESC2.setBounds(125,390,300,25);
        DESC3 = new JTextField();
        DESC3.setBounds(125,425,300,25);
        DESC4 = new JTextField();
        DESC4.setBounds(125,460,300,25);
        DESC5 = new JTextField();
        DESC5.setBounds(125,495,300,25);
        DESC6 = new JTextField();
        DESC6.setBounds(125,530,300,25);
        DESC7 = new JTextField();
        DESC7.setBounds(125,565,300,25);
        DESC8 = new JTextField();
        DESC8.setBounds(125,600,300,25);
        DESC9 = new JTextField();
        DESC9.setBounds(125,635,300,25);
        panel.add(DESC);
        panel.add(DESC1);
        panel.add(DESC2);
        panel.add(DESC3);
        panel.add(DESC4);
        panel.add(DESC5);
        panel.add(DESC6);
        panel.add(DESC7);
        panel.add(DESC8);
        panel.add(DESC9);
    }

    private void botonPrint(){
        botonPrint = new JButton();
        botonPrint.setBounds(310, 110,120,100);
        botonPrint.setText("Print");
        botonPrint.addActionListener(null);
        ActionListener print = new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    //Llamamos al archivo
                    FileInputStream file = new FileInputStream(new File("Excell.xlsx"));

                    //Elegimos el archivo sobre el que operaremos
                    XSSFWorkbook wb = new XSSFWorkbook(file);

                    //Elegimos la hoja sobre la que operaremos
                    XSSFSheet hoja = wb.getSheetAt(0);

                    //Declarando las filas a utilizar
                    Row Fila4 = hoja.getRow(3);
                    Row Fila12 = hoja.getRow(11);
                    Row Fila13 = hoja.getRow(12);
                    Row Fila16 = hoja.getRow(15);
                    Row Fila17 = hoja.getRow(16);
                    Row Fila18 = hoja.getRow(17);
                    Row Fila19 = hoja.getRow(18);
                    Row Fila20 = hoja.getRow(19);
                    Row Fila21 = hoja.getRow(20);
                    Row Fila22 = hoja.getRow(21);
                    Row Fila23 = hoja.getRow(22);
                    Row Fila24 = hoja.getRow(23);
                    Row Fila25 = hoja.getRow(24);
                    Row Fila50 = hoja.getRow(49);
                    Row Fila51 = hoja.getRow(50);


                    //Declaramos las Celdas/Columnas
                    //Fecha
                    Cell Dia = Fila4.getCell(9);
                    Cell Mes = Fila4.getCell(10);
                    Cell Año = Fila4.getCell(11);

                    //Rellenando
                    Dia.setCellValue(DIA.getText());
                    Mes.setCellValue(MES.getText());
                    Año.setCellValue(AÑO.getText());

                    //Datos del destinatario
                    Cell Nombre = Fila12.getCell(2);
                    Cell Cuit = Fila12.getCell(7);
                    Cell Dirección = Fila13.getCell(2);
                    Cell Responsable = Fila13.getCell(7);

                    //Rellenando
                    Nombre.setCellValue(NOMBRE.getText());
                    Cuit.setCellValue(CUIT.getText());
                    Dirección.setCellValue(DOMICILIO.getText());
                    Responsable.setCellValue(CONDICION.getText());

                    //Datos del trasportista
                    Cell ResSoc = Fila50.getCell(2);
                    Cell CuitT = Fila50.getCell(7);
                    Cell NombreT = Fila51.getCell(2);
                    Cell Dni = Fila51.getCell(7);

                    //Rellenando
                    ResSoc.setCellValue(RAZON.getText());
                    CuitT.setCellValue(CUITT.getText());
                    NombreT.setCellValue(NOMBRET.getText());
                    Dni.setCellValue(DNI.getText());

                    //Cant Ítems
                    Cell Cant = Fila16.getCell(1);
                    Cell Cant1 = Fila17.getCell(1);
                    Cell Cant2 = Fila18.getCell(1);
                    Cell Cant3 = Fila19.getCell(1);
                    Cell Cant4 = Fila20.getCell(1);
                    Cell Cant5 = Fila21.getCell(1);
                    Cell Cant6 = Fila22.getCell(1);
                    Cell Cant7 = Fila23.getCell(1);
                    Cell Cant8 = Fila24.getCell(1);
                    Cell Cant9 = Fila25.getCell(1);

                    //Rellenando
                    Cant.setCellValue(CANT.getText());
                    Cant1.setCellValue(CANT1.getText());
                    Cant2.setCellValue(CANT2.getText());
                    Cant3.setCellValue(CANT3.getText());
                    Cant4.setCellValue(CANT4.getText());
                    Cant5.setCellValue(CANT5.getText());
                    Cant6.setCellValue(CANT6.getText());
                    Cant7.setCellValue(CANT7.getText());
                    Cant8.setCellValue(CANT8.getText());
                    Cant9.setCellValue(CANT9.getText());

                    //Unidad Ítems
                    Cell Un = Fila16.getCell(2);
                    Cell Un1 = Fila17.getCell(2);
                    Cell Un2 = Fila18.getCell(2);
                    Cell Un3 = Fila19.getCell(2);
                    Cell Un4 = Fila20.getCell(2);
                    Cell Un5 = Fila21.getCell(2);
                    Cell Un6 = Fila22.getCell(2);
                    Cell Un7 = Fila23.getCell(2);
                    Cell Un8 = Fila24.getCell(2);
                    Cell Un9 = Fila25.getCell(2);

                    //Rellenando
                    Un.setCellValue(UN.getText());
                    Un1.setCellValue(UN1.getText());
                    Un2.setCellValue(UN2.getText());
                    Un3.setCellValue(UN3.getText());
                    Un4.setCellValue(UN4.getText());
                    Un5.setCellValue(UN5.getText());
                    Un6.setCellValue(UN6.getText());
                    Un7.setCellValue(UN7.getText());
                    Un8.setCellValue(UN8.getText());
                    Un9.setCellValue(UN9.getText());

                    //Descripción Ítems
                    Cell Desc = Fila16.getCell(3);
                    Cell Desc1 = Fila17.getCell(3);
                    Cell Desc2 = Fila18.getCell(3);
                    Cell Desc3 = Fila19.getCell(3);
                    Cell Desc4 = Fila20.getCell(3);
                    Cell Desc5 = Fila21.getCell(3);
                    Cell Desc6 = Fila22.getCell(3);
                    Cell Desc7 = Fila23.getCell(3);
                    Cell Desc8 = Fila24.getCell(3);
                    Cell Desc9 = Fila25.getCell(3);

                    //Rellenando
                    Desc.setCellValue(DESC.getText());
                    Desc1.setCellValue(DESC1.getText());
                    Desc2.setCellValue(DESC2.getText());
                    Desc3.setCellValue(DESC3.getText());
                    Desc4.setCellValue(DESC4.getText());
                    Desc5.setCellValue(DESC5.getText());
                    Desc6.setCellValue(DESC6.getText());
                    Desc7.setCellValue(DESC7.getText());
                    Desc8.setCellValue(DESC8.getText());
                    Desc9.setCellValue(DESC9.getText());

                    FileOutputStream fileout = new FileOutputStream("Excell.xlsx");
                    wb.write(fileout);
                    fileout.close();
                    File fileToPrint = new File("Excell.xlsx");
                    Desktop.getDesktop().print(fileToPrint);
                    Desktop.getDesktop().print(fileToPrint);
                    Desktop.getDesktop().print(fileToPrint);

                    JOptionPane.showMessageDialog(null, "Se envío el trabajo a la impresora predeterminada");

                } catch (FileNotFoundException f) {
                    throw new RuntimeException(f);
                } catch (IOException f) {
                    throw new RuntimeException(f);
                }
            }
        };
        botonPrint.addActionListener(print);
        panel.add(botonPrint);

    }
    private void botonExit(){
        botonExit = new JButton();
        botonExit.setBounds(115,665,250,30);
        botonExit.setText("Exit");
        ActionListener Exit = new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                dispose();
            }
        };
        botonExit.addActionListener(Exit);
        panel.add(botonExit);
    }
    private void Botones(){
        botonPrint();
        botonExit();
    }

}


