  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPrincipal = planilla.getSheetByName('Calculadora calorías');






class Alimento 
{
  constructor(nombre, calorias, proteina, carbohidratos, grasa) 
  {
    if (this.constructor === Alimento) 
      {
        throw new Error("Alimento es una clase abstracta");
      }

    this.nombre = nombre;
    this.calorias = calorias;
    this.proteina = proteina;
    this.carbohidratos = carbohidratos;
    this.grasa = grasa;
  }

  getMacros(gramos) 
  {
    return {
      calorias: this.calorias * gramos,
      proteina: this.proteina * gramos,
      carbohidratos: this.carbohidratos * gramos,
      grasa: this.grasa * gramos
    };
  }
}







// A L I M E N T O S
let huevoEntero = {};
huevoEntero.nombre = "Huevo entero";
huevoEntero.proteina = 0.126;
huevoEntero.grasa = 0.099;
huevoEntero.carbohidratos = 0.008;
huevoEntero.calorias = 1.467;

let nueces = {};
nueces.nombre = "Nueces";
nueces.proteina = 0.04;
nueces.grasa = 0.16;
nueces.carbohidratos = 0.01;
nueces.calorias = 1.64;

let almendras = {};
almendras.nombre = "Almendras";
almendras.proteina = 0.21;
almendras.grasa = 0.56;
almendras.carbohidratos = 0.04;
almendras.calorias = 6.28;

let mani = {};
mani.nombre = "Maní";
mani.proteina = 0.26;
mani.grasa = 0.5;
mani.carbohidratos = 0.16;
mani.calorias = 6.18;

let proteina = {};
mani.nombre = "proteina";
mani.proteina = 0.8;
mani.grasa = 0.05;
mani.carbohidratos = 0.14;
mani.calorias = 3.86;


let pataYMuslo =
{
  nombre: "Pata y muslo",
  proteina: 0.201,
  grasa: 0.038,
  carbohidratos: 0,
  calorias: 1.2
}

let papas =
{
  nombre: "Papas",
  proteina: 0.01,
  grasa: 0,
  carbohidratos: 0.18,
  calorias: 0.76
}

let palta =
{
  nombre: "Palta",
  proteina: 0.022,
  grasa: 0.156,
  carbohidratos: 0.078,
  calorias: 1.8
}

let jamonCocido =
{
  nombre: "Jamon cocido",
  proteina: 0.189,
  grasa: 0.032,
  carbohidratos: 0.012,
  calorias: 1.09
}

let pechugaPollo =
{
  nombre: "Pechuga de pollo",
  proteina: 0.22,
  grasa: 0.06,
  carbohidratos: 0,
  calorias: 1.45
}

let roastBeef =
{
  nombre: "Roast Beef",
  proteina: 0.22,
  grasa: 0.03,
  carbohidratos: 0,
  calorias: 1.2
}

let arroz =
{
  nombre: "Arroz",
  proteina: 0.3,
  grasa: 0,
  carbohidratos: 0.43,
  calorias: 1.8
}

let avena =
{
  nombre: "Avena",
  proteina: 0.14,
  grasa: 0,
  carbohidratos: 0.57,
  calorias: 3.7
}

let manteca = {};
manteca.nombre = "Manteca";
manteca.proteina = 0;
manteca.grasa = 0.8;
manteca.carbohidratos = 0;
manteca.calorias = 7.6;

let carnePicadaRes =
{
  nombre: "Carne picada de res",
  proteina: 0.18,
  grasa: 0.06,
  carbohidratos: 0.03,
  calorias: 1.37
}

let tapaDeNalga =
{
  nombre: "Tapa de nalga",
  proteina: 0.22,
  grasa: 0.08,
  carbohidratos: 0.00,
  calorias: 1.57
}

let filetMerluza =
{
  nombre: "Filet de merluza",
  proteina: 0.12,
  grasa: 0.02,
  carbohidratos: 0.00,
  calorias: 0.65
}

let aceiteDeOliva =
{
  nombre: "Aceite de oliva",
  proteina: 0,
  grasa: 5,
  carbohidratos: 0,
  calorias: 45
}

let lomoDeAtun =
{
  nombre: "Lomo de atún",
  proteina: 0.2,
  grasa: 0.1,
  carbohidratos: 0.1,
  calorias: 0.89
}

let caballa =
{
  nombre: "Caballa",
  proteina: 0.21,
  grasa: 0.18,
  carbohidratos: 0,
  calorias: 2.45
}


// R E C E T A S
let polloMostaza =
{
  nombre: "Pollo a la mostaza",
  alimentos: [pataYMuslo, papas],
  cantidades: [100, 50],
  url: "file:///F:/Users/Francisco/Documents/Documentos/Archivos/Libros%20y%20artículos/FitnessRevolucionario/Plan%20R28%20-%20Fitness%20Revolucionario.pdf#page=108"
}

let huevosRevueltos =
{
  nombre: "Huevos revueltos",
  alimentos: [huevoEntero, jamonCocido, palta],
  cantidades: [195, 30, 70],
  url: "file:///F:/Users/Francisco/Documents/Documentos/Archivos/Libros%20y%20artículos/FitnessRevolucionario/Plan%20R28%20-%20Fitness%20Revolucionario.pdf#page=50"
}


let carneAlHorno = {};
carneAlHorno.nombre = "Carne al horno";
carneAlHorno.alimentos = [roastBeef, papas];
carneAlHorno.cantidades = [700, 1000];

let pureDePapa = {};
pureDePapa.nombre = "Puré de papa";
pureDePapa.alimentos = [papas, manteca];
pureDePapa.cantidades = [500, 50];

let filetMerluzaAlHorno = {};
filetMerluzaAlHorno.nombre = "Filet de merluza al horno";
filetMerluzaAlHorno.alimentos = [filetMerluza, papas];
filetMerluzaAlHorno.cantidades = [1200, 700];

let pastelDePapasConEnsalada = {};
pastelDePapasConEnsalada.nombre = "Pastel de papas con ensalada";
pastelDePapasConEnsalada.alimentos = [papas, carnePicadaRes, aceiteDeOliva];
pastelDePapasConEnsalada.cantidades = [2000, 1000, 3];

let ensaladaEspartana = {};
ensaladaEspartana.nombre = "Ensalada espartana";
ensaladaEspartana.alimentos = [lomoDeAtun, aceiteDeOliva];
ensaladaEspartana.cantidades = [150, 0.5];

let polloAlLimonConPureDePapa = {};
polloAlLimonConPureDePapa.nombre = "Pollo al limón con puré de papa";
polloAlLimonConPureDePapa.alimentos = [pechugaPollo, papas, manteca];
polloAlLimonConPureDePapa.cantidades = [200, 500, 50];

// LISTAS
var alimentos = [roastBeef, papas, palta, jamonCocido, pataYMuslo, huevoEntero, pechugaPollo, arroz, avena, manteca, carnePicadaRes, aceiteDeOliva, lomoDeAtun, caballa, tapaDeNalga, filetMerluza, nueces, almendras, mani, proteina];

var recetas = [carneAlHorno, huevosRevueltos, polloMostaza, pureDePapa, pastelDePapasConEnsalada, ensaladaEspartana, polloAlLimonConPureDePapa, filetMerluzaAlHorno];


// CELDAS A EDITAR
  var breakfastRecipeCell = hojaPrincipal.getRange(2,2);
  var lunchReciteCell = hojaPrincipal.getRange(10,2);
  var dinnerRecipeCell = hojaPrincipal.getRange(18,2);
  var recipeCellGroup = [breakfastRecipeCell, lunchReciteCell, dinnerRecipeCell];

  var breakfastPortionCells = hojaPrincipal.getRange(3,2);
  var lunchPortionCells = hojaPrincipal.getRange(11,2);
  var dinnerPortionCells = hojaPrincipal.getRange(19,2);
  var portionRecipesCellGroup = [breakfastPortionCells, lunchPortionCells, dinnerPortionCells];


  var foodCells = hojaPrincipal.getRange(2,3,1,8);
  var foodCells2 = hojaPrincipal.getRange(10,3,1,8);
  var foodCells3 = hojaPrincipal.getRange(18,3,1,8);
  var foodCellsGroup = [foodCells, foodCells2, foodCells3];

  var cellAmountOfFood1 = hojaPrincipal.getRange(3,3,1,8);
  var cellAmountOfFood2 = hojaPrincipal.getRange(11,3,1,8);
  var cellAmountOfFood3 = hojaPrincipal.getRange(19,3,1,8);
  var groupOfCellsAmountOfFood = [cellAmountOfFood1, cellAmountOfFood2, cellAmountOfFood3];




function onEdit(e)
{
  var editedRange = e.range; // Obtener el rango editado


  if( devolverCelda_SiCoincideConConjunto_(editedRange, foodCellsGroup) )
  {
    aporteAlimento(editedRange);
  }  
  else if( devolverCelda_SiCoincideConConjunto_(editedRange, groupOfCellsAmountOfFood) )
  {
    aporteAlimento( hojaPrincipal.getRange(editedRange.getRow()-1, editedRange.getColumn()) );
  }
  else if( devolverCelda_SiCoincideConCeldaDeArray_(editedRange, recipeCellGroup) )
  {
    accionarReceta(editedRange);
  }
  else if( devolverCelda_SiCoincideConCeldaDeArray_(editedRange, portionRecipesCellGroup) )
  {
    accionarReceta( hojaPrincipal.getRange(editedRange.getRow()-1, editedRange.getColumn()) );
  }

}


function program()
{
  var celdaAsignada = hojaPrincipal.getRange(18,6);

  Logger.log( aporteAlimento(celdaAsignada) );
}


function devolverCelda_SiCoincideConCeldaDeArray_(celda, array)
{
  var celdaCoincidente = false;
  var copiaArray = array.slice();

  while( !celdaCoincidente && copiaArray.length != 0 )
  {
    if( celda.getA1Notation() == copiaArray.shift().getA1Notation() )
    {
      celdaCoincidente = true;
    }
  }
  return celdaCoincidente;
}


function devolverCelda_SiCoincideConConjunto_(celda, conjuntoGrupoCeldas)
/*
  Propósito: Retorna 'celda' si la misma coincide con alguna de las celdas contenidas en 'conjuntoGrupoCeldas'.
  
  Parámetros: 'celda'. Tipo: Range. La celda a contrastar con el grupo de celdas
    'conjuntoGrupoCeldas'. Tipo: Lista de range.
*/
{
  var celdaCoincidente = false;
  var copiaConjuntoCeldas = conjuntoGrupoCeldas.slice();

  while( !celdaCoincidente && copiaConjuntoCeldas.length != 0 )
  {
    celdaCoincidente = celda_CoincideConCeldaDeGrupo_(celda, copiaConjuntoCeldas.shift() );
  }
  return celdaCoincidente;
}


function celda_CoincideConCeldaDeGrupo_(celda, grupoCeldas)
{
  var celdaCoincidente = false;

  for( var columnas = 1; columnas <= 8; columnas++ )
  {
    if( celda.getA1Notation() == grupoCeldas.getCell(1, columnas).getA1Notation() )
    {
      celdaCoincidente = true;
    } 
  }
  return celdaCoincidente;
}


function accionarReceta(celdaReceta)
{
  var celdaRecetaSelect = celdaReceta.getValue();
  var recetaSeleccionada = posicionReceta_(celdaRecetaSelect);
  var comidas = [ [] ];
  var cantidad = [ [] ];
  var listaAlimentos = recetas[recetaSeleccionada].alimentos;
  var listaCantidades = recetas[recetaSeleccionada].cantidades;
  var porcionDeLaReceta = hojaPrincipal.getRange(celdaReceta.getRow()+1, celdaReceta.getColumn()).getValue();


  for( var elementos = 0; elementos<listaAlimentos.length; elementos++ )
  {
    comidas[0].push(listaAlimentos[elementos].nombre);
    cantidad[0].push(listaCantidades[elementos] * porcionDeLaReceta);
  }


  // Establece los alimentos de la receta
  hojaPrincipal.getRange(celdaReceta.getRow(), celdaReceta.getColumn()+1, 1, elementos).setValues(comidas);

  // Establece las cantidades de esos alimentos de la receta
  hojaPrincipal.getRange(celdaReceta.getRow()+1, celdaReceta.getColumn()+1, 1, elementos).setValues(cantidad);

  // Elimina los alimentos sobrantes de recetas accionadas anteriormente
  hojaPrincipal.getRange(celdaReceta.getRow(), celdaReceta.getColumn()+elementos+1, 6, 8-elementos).clearContent();
  
  // Aporte nutricional de los respectivos alimentos de la receta seleccionada
  for( var comidas = 1; comidas <= recetas[recetaSeleccionada].alimentos.length; comidas++ )
  {
    aporteAlimento( hojaPrincipal.getRange(celdaReceta.getRow(), celdaReceta.getColumn() + comidas) );
  }
}


function aporteAlimento(celdaAlimento)
{
  var alimentoSeleccionado = posicionAlimento_( celdaAlimento.getValue() );
 
  var cantidad = hojaPrincipal.getRange( celdaAlimento.getRow()+1, celdaAlimento.getColumn() ).getValue();

  var aporteNutricional = [ [ alimentos[alimentoSeleccionado].proteina * cantidad], [alimentos[alimentoSeleccionado].carbohidratos * cantidad], [alimentos[alimentoSeleccionado].grasa * cantidad], [alimentos[alimentoSeleccionado].calorias * cantidad] ];
  
  hojaPrincipal.getRange(celdaAlimento.getRow()+2, celdaAlimento.getColumn(), 4, 1).setValues(aporteNutricional);
}


function posicionAlimento_(nombreAlimento)
{
  return alimentos.findIndex(alimento => alimento.nombre == nombreAlimento);
}


function posicionReceta_(nombreReceta)
{
  return recetas.findIndex(elemento => elemento.nombre == nombreReceta);
}





