import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { CommonModule } from '@angular/common';
import { saveAs } from 'file-saver';
import { Workbook } from 'exceljs';

@Component({
  selector: 'app-cotizador',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './cotizador.component.html',
  styleUrls: ['./cotizador.component.scss']
})
export class CotizadorComponent {
  datosCombinados: any[] = [];
  archivoBOM: File | null = null;
  archivoOferta: File | null = null;
  errorMsg: string = '';

  onFileChange(event: any, tipo: string): void {
    const archivo = event.target.files[0];
    if (tipo === 'bom') {
      this.archivoBOM = archivo;
    } else {
      this.archivoOferta = archivo;
    }
  }

  procesarArchivos(): void {
    this.errorMsg = '';
    if (!this.archivoBOM || !this.archivoOferta) {
      this.errorMsg = 'Debe subir ambos archivos.';
      return;
    }

    const lectorBOM = new FileReader();
    const lectorOferta = new FileReader();

    lectorBOM.onload = (e: any) => {
      const workbookBOM = XLSX.read(e.target.result, { type: 'binary' });
      const hojaBOM = workbookBOM.Sheets[workbookBOM.SheetNames[0]];
      const datosBOM = XLSX.utils.sheet_to_json(hojaBOM);

      lectorOferta.onload = (e2: any) => {
        const workbookOferta = XLSX.read(e2.target.result, { type: 'binary' });
        const hojaOferta = workbookOferta.Sheets["OFERTA DEL DISTRIBUIDOR"];

        let datosOferta = XLSX.utils.sheet_to_json(hojaOferta, { header: 1 }) as any[][];
        const headerRowIndex = datosOferta.findIndex((fila: any[]) =>
          fila.some((cell: any) => String(cell).toLowerCase().includes('mfr. part'))
        );

        if (headerRowIndex === -1) {
          this.errorMsg = 'No se encontró encabezado "Mfr. Part #" en la Oferta.';
          return;
        }

        const headers = datosOferta[headerRowIndex];
        const filas = datosOferta.slice(headerRowIndex + 1);
        const ofertaDataFormatted = filas.map((fila: any[]) => {
          const obj: any = {};
          headers.forEach((header: string, i: number) => obj[header] = fila[i]);
          return obj;
        });

        this.combinarDatos(datosBOM, ofertaDataFormatted);
      };

      if (this.archivoOferta) lectorOferta.readAsBinaryString(this.archivoOferta);
    };

    if (this.archivoBOM) lectorBOM.readAsBinaryString(this.archivoBOM);
  }

  private combinarDatos(bom: any[], oferta: any[]): void {
    const colBOM = Object.keys(bom[0]).find(c => c.toLowerCase().includes('part number'));
    const colOferta = Object.keys(oferta[0]).find(c => c.toLowerCase().includes('mfr. part'));

    if (!colBOM || !colOferta) {
      this.errorMsg = 'No se encontraron columnas Part Number o Mfr. Part #.';
      return;
    }

    this.datosCombinados = bom.map(itemBOM => {
      const match = oferta.find(itemOferta =>
        String(itemOferta[colOferta]).trim() === String(itemBOM[colBOM]).trim()
      );
      return { ...itemBOM, ...match };
    }).filter(item => Object.keys(item).length > Object.keys(bom[0]).length);
  }

  getColumnKeys(): string[] {
    if (this.datosCombinados.length === 0) return [];
    return Object.keys(this.datosCombinados[0])
      .filter(k => k !== 'Part Number');
  }

  exportarExcel(): void {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet('Cotizacion');


  const headers = Object.keys(this.datosCombinados[0]);
  worksheet.addRow(headers);

  
  this.datosCombinados.forEach(item => {
    worksheet.addRow(Object.values(item));
  });

  
  worksheet.getRow(1).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2D456E' } };
    cell.alignment = { horizontal: 'center' };
  });

  workbook.xlsx.writeBuffer().then(buffer => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'resultado_cotizacion.xlsx');
  });
}

  imprimirTabla(): void {
    const contenido = document.getElementById('tablaCotizacion')?.innerHTML;
    if (!contenido) {
      alert('No hay datos para imprimir');
      return;
    }

    
    const ventana = window.open('', '_blank', 'width=1000,height=600');

    if (ventana) {
      ventana.document.open();
      ventana.document.write(`
      <html>
        <head>
          <title>Imprimir Cotización</title>
          <style>
            table { width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #001E43; color: white; }
          </style>
        </head>
        <body onload="window.print(); window.close();">
          ${contenido}
        </body>
      </html>
    `);
      ventana.document.close();
    } else {
      alert('El navegador bloqueó la ventana emergente. Permite pop-ups e inténtalo de nuevo.');
    }
  }


}


