import XLSX from "xlsx"

function limparNumero(valor) {
  if (!valor) return ""
  return String(valor).replace(/\D/g, "")
}

function validarCPF(cpf) {
  cpf = limparNumero(cpf)

  if (cpf.length !== 11) return false

  if (/^(\d)\1{10}$/.test(cpf)) return false

  let soma = 0
  for (let i = 0; i < 9; i++) {
    soma += parseInt(cpf[i]) * (10 - i)
  }
  let resto = (soma * 10) % 11
  if (resto === 10 || resto === 11) resto = 0
  if (resto !== parseInt(cpf[9])) return false

  soma = 0
  for (let i = 0; i < 10; i++) {
    soma += parseInt(cpf[i]) * (11 - i)
  }
  resto = (soma * 10) % 11
  if (resto === 10 || resto === 11) resto = 0
  if (resto !== parseInt(cpf[10])) return false

  return true
}

function validarCNPJ(cnpj) {
  cnpj = limparNumero(cnpj)

  if (cnpj.length !== 14) return false

  if (/^(\d)\1{13}$/.test(cnpj)) return false

  const mult1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
  const mult2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

  // Calcula o primeiro dígito verificador
  let soma = 0
  for (let i = 0; i < 12; i++) {
    soma += parseInt(cnpj[i]) * mult1[i]
  }
  let resto = soma % 11
  const dig1 = resto < 2 ? 0 : 11 - resto
  if (dig1 !== parseInt(cnpj[12])) return false

  soma = 0
  for (let i = 0; i < 13; i++) {
    soma += parseInt(cnpj[i]) * mult2[i]
  }
  resto = soma % 11
  const dig2 = resto < 2 ? 0 : 11 - resto
  if (dig2 !== parseInt(cnpj[13])) return false

  return true
}

function validarDocumento(valor) {
  const numero = limparNumero(valor)

  if (numero.length === 11) {
    return { tipo: "CPF", valido: validarCPF(numero) }
  } else if (numero.length === 14) {
    return { tipo: "CNPJ", valido: validarCNPJ(numero) }
  } else {
    return { tipo: numero.length < 14 ? "CPF" : "CNPJ", valido: false }
  }
}

function main() {
  console.log("=".repeat(50))
  console.log("VALIDADOR DE CPF/CNPJ - CLIENTES")
  console.log("=".repeat(50))
  console.log()

  const workbook = XLSX.readFile("clientes.xlsx")
  const sheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[sheetName]

  const dados = XLSX.utils.sheet_to_json(worksheet)

  const clientesInvalidos = []
  let cpfInvalidos = 0
  let cnpjInvalidos = 0

  dados.forEach((cliente) => {
    const documento = cliente["CPF/CNPJ"]
    const resultado = validarDocumento(documento)

    if (!resultado.valido) {
      clientesInvalidos.push(cliente)

      if (resultado.tipo === "CPF") {
        cpfInvalidos++
      } else {
        cnpjInvalidos++
      }
    }
  })

  console.log("RESULTADO DA VALIDAÇÃO:")
  console.log("-".repeat(50))
  console.log(`Total de registros processados: ${dados.length}`)
  console.log(`Total de CPF/CNPJ inválidos: ${clientesInvalidos.length}`)
  console.log(`  - CPFs inválidos: ${cpfInvalidos}`)
  console.log(`  - CNPJs inválidos: ${cnpjInvalidos}`)
  console.log("-".repeat(50))

  if (clientesInvalidos.length > 0) {
    const novaWorksheet = XLSX.utils.json_to_sheet(clientesInvalidos)
    const novoWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(
      novoWorkbook,
      novaWorksheet,
      "Clientes Inválidos"
    )

    XLSX.writeFile(novoWorkbook, "clientes_invalidos.xlsx")

    console.log()
    console.log('✓ Arquivo "clientes_invalidos.xlsx" gerado com sucesso!')
    console.log(
      `  Contém ${clientesInvalidos.length} registros com CPF/CNPJ inválido.`
    )
  } else {
    console.log()
    console.log("✓ Todos os CPF/CNPJ são válidos! Nenhum arquivo gerado.")
  }

  console.log()
  console.log("=".repeat(50))
}

main()
