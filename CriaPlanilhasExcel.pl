use warnings;
use strict;

use Excel::Writer::XLSX;

# Define variaveis globais
my @Meses = ( 'NONE', 'JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO' );
my @DiasNoMes = ( 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

sub CriarPasta {
  
  # recebe data a ser usada, em formato dd/mm/yyyy
  # separa mes
  my ( $Data ) = @_;
  my $MesStr = substr $Data, 3, 2;
    
  # mostra mes a ser criado
  my $Pasta = $MesStr . " " . $Meses[$MesStr];
  print "Pasta: ", $Pasta, "\n";
  
  # cria pasta
  if ( mkdir ( $Pasta ) ) {
    print "Pasta '$Pasta' criada com sucesso.\n";
  } else {
    print "Erro ao criar pasta '$Pasta': $!\n";
    exit;
  }
}

sub CriarWorkBook {
  
  # Parametros: mes a ser criado e data a ser colocada
  my ( $Data ) = @_;
  
  # Trata o mes para criar o nome do arquivo no formato
  # "PLANILHA-01-JANEIRO.xlsx"
  my $DiaStr = substr ( $Data, 0, 2);
  my $MesStr = substr ( $Data, 3, 2);
  
  # preparar para criar o WorkBook
  my $Pasta   = $MesStr . " " . $Meses[$MesStr];
  my $Arquivo =  "PLANILHA " . $Meses[$MesStr] . " " . $DiaStr .  ".xlsx";

  # Entra na pasta e cria planilha
  if ( !( chdir ($Pasta) ) ) {
    print "Pasta '$Pasta' parece nao existir.\n";
    exit;
  }

  # Create a new Excel workbook
  my $WorkBook = Excel::Writer::XLSX->new( $Arquivo );
  
  # Adiciona a WorkSheet e os formatos
  my $WorkSheet       = $WorkBook->add_worksheet();
  my $FormatCabecalho = $WorkBook->add_format();
  my $FormatCenter    = $WorkBook->add_format();
  my $FormatoLeft     = $WorkBook->add_format();

  # Define formato dos cabecalhos
  $FormatCabecalho->set_bold();
  $FormatCabecalho->set_size( 16 );
  $FormatCabecalho->set_bg_color( 'yellow' );
  $FormatCabecalho->set_align( 'center' );
  $FormatCabecalho->set_border ( 1 );

  # Define formato das colunas com alinhamento left
  $FormatoLeft->set_border ( 1 );
  $FormatoLeft->set_align ( 'left' );

  # Define formato das colunas B e C
  $FormatCenter->set_border ( 1 );
  $FormatCenter->set_align ( 'center' );

  # Formata tamanho das colunas
  $WorkSheet->set_column ( 0, 0, 12 );
  $WorkSheet->set_column ( 1, 2, 40 );
  $WorkSheet->set_column ( 3, 6, 15 );

  # Coloca cabecalhos
  $WorkSheet->write ( "A1", "DATA",    $FormatCabecalho );
  $WorkSheet->write ( "B1", "NOME",    $FormatCabecalho );
  $WorkSheet->write ( "C1", "EMPRESA", $FormatCabecalho );
  $WorkSheet->write ( "D1", "ENTRADA", $FormatCabecalho );
  $WorkSheet->write ( "E1", "PLACA",   $FormatCabecalho );
  $WorkSheet->write ( "F1", "MODELO",  $FormatCabecalho );
  $WorkSheet->write ( "G1", "SAIDA",   $FormatCabecalho );

  # Coloca linhas
  for ( my $lin = 1; $lin < 100; $lin++ ) {
    $WorkSheet->write_string ( $lin, 0, $Data, $FormatCenter );
    $WorkSheet->write_string ( $lin, 1, '', $FormatoLeft );
    $WorkSheet->write_string ( $lin, 2, '', $FormatoLeft );
    $WorkSheet->write_string ( $lin, 3, '', $FormatCenter );
    $WorkSheet->write_string ( $lin, 4, '', $FormatoLeft );
    $WorkSheet->write_string ( $lin, 5, '', $FormatoLeft );
    $WorkSheet->write_string ( $lin, 6, '', $FormatCenter ); 
  }

  # fecha a planilha
  $WorkBook->close();
  
  # sai da pasta
  chdir ( ".." );
  
  # mensagem
  print "Planilha $Arquivo criada com sucesso.\n";
}

for ( my $MesDoAno = 1; $MesDoAno <= 12; $MesDoAno++ ) {
  
  # Prepara para criar a pasta
  my $MesDaPlanilha = substr ( "000".$MesDoAno, -2 );
  
  # Cria a pasta
  CriarPasta ( "01/" . $MesDaPlanilha . "/2026" );
  
  for ( my $DiaDoMes = 1; $DiaDoMes <= $DiasNoMes[$MesDoAno]; $DiaDoMes++ ) {
    # Prepara para criar planilha
    my $DiaDaPlanilha = substr ( "000".$DiaDoMes, -2 );
    
    CriarWorkBook ( $DiaDaPlanilha . "/" . $MesDaPlanilha . "/2026" );
  }
  
  print ( "\n" );
}

exit;
#
# EOF
#