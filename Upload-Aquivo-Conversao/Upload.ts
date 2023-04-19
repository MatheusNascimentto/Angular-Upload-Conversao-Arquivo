 private converterXls(arquivoExcel: any) {
    let fileReader = new FileReader();
    fileReader.readAsArrayBuffer(arquivoExcel);
    fileReader.onload = (e) => {
      let arrayBuffer = fileReader.result;
      const workbook = XLSX.read(arrayBuffer, { type: !!fileReader.readAsBinaryString ? 'binary' : 'array' });
      let sheet_name = workbook.SheetNames[0];
      let ws = workbook.Sheets[sheet_name];
      this.delete_row(ws, 0);
      this.delete_row(ws, 1);
      let resultado = XLSX.utils.sheet_to_json(ws, { raw: true, header: ['personId', 'firstName', 'lastName', 'username', 'country', 'company', 'location', 'employeeStatus', 'employeeClass', 'startDate', 'gender', 'dateOfBirth', 'nationality', 'jobClassification', 'costCentreCode', 'costCentre', 'jobFunction', 'jobLevel', 'fte', 'businessUnit', 'division', 'department', 'contractType', 'contractEndDate', 'jobTitle', 'emailAddress', 'emailType', 'isPrimary', 'lineManager', 'lineManagerFirstName', 'lineManagerLastName', 'workSchedule', 'holidayCalendar', 'timeProfile', 'timeOffManager', 'timeOffManagerFirstName', 'timeOffManagerLastName', 'jobFunctionCode'], skipHidden: true, defval: '' })
      resultado.forEach(r => {
        let t: Table = r as Table
        let tabela: Tabela = new Tabela(t);
        this.colaboradoresTabela.push(tabela);
      })
      this.mensagemService.add("Cadastrando colaboradores... Aguarde!");
    }
  }

  private ec(r: any, c: any) {
    return XLSX.utils.encode_cell({ r: r, c: c });
  }

  private delete_row(ws: any, row_index: any) {
    var variable = XLSX.utils.decode_range(ws["!ref"])
    for (var R = row_index; R < variable.e.r; ++R) {
      for (var C = variable.s.c; C <= variable.e.c; ++C) {
        ws[this.ec(R, C)] = ws[this.ec(R + 1, C)];
      }
    }
    variable.e.r--
    ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
  }

  // É necessario instalar uma biblioteca do XLSX