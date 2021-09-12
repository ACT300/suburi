/* 利用ファイルのヘッダ */
#include "comp_func.h"
#include "decomp_func.h"

/* 自ファイルのヘッダ */
#include "main.h"

static void printMenu(void);
static void selectCompFunc(comp_t*);
static void selectDecompFunc(comp_t*);
static void releaseCompParams(comp_t*);

int main(void){
  char input[] = "/home/vagrant/projects/suburi/compression/comp/bin.bin";
  char output[] = "/home/vagrant/projects/suburi/compression/comp/result";
  FILE *fp_i;
  FILE *fp_o;
  int datasize;
  int c;
  unsigned char buffer[BUFFER_SIZE];
  int val;
  char str[8];
  comp_t compParams;

  printf("モードを選択してください。\n");
  printf("[1]圧縮モード\n");
  printf("[2]伸張モード\n");
  printf("[3]データベースを見る\n");

  switch(mode){
    case 1:
      break;
    case 2:
      break;
    case 3:
      break;
    defalt:
      break;  
  }

  printMenu();

  while(1){
    do{
      fgets(str, sizeof(str), stdin);
      val = atoi(str);

      if(val <= NOT_SELECT_RUNLENGTH || RUNLENGTH_NUM <= val){
        printf("不正な値です。\n");
      }
    }while(val <= NOT_SELECT_RUNLENGTH || RUNLENGTH_NUM <= val);

    fp_i = fopen(input, "rb");
    if(fp_i == NULL){
      fprintf(stderr, "Input File cannot opne.\n");
      exit(8);
    }

    datasize = 0;
    while(datasize < BUFFER_SIZE && (c=fgetc(fp_i)) != EOF){
      buffer[datasize] = c;
      datasize++;
    }
    fclose(fp_i);

    fp_o = fopen(output, "wb");
    if(fp_o == NULL){
      fprintf(stderr, "Output File cannot open.\n");
      exit(8);
    }

    compParams.fp_o = fp_o;
    compParams.data = buffer;
    compParams.size = datasize;
    compParams.event = val;

    selectCompFunc(&compParams);

    printf("符号化に成功しました。 符号化に使用した関数番号[%d], 元データのサイズ[%d]Byte\n", compParams.event, compParams.size);

    fclose(fp_o);
    releaseCompParams(&compParams);
  }
}

static void printMenu(void){
  printf("-------------------圧縮・伸張プログラム-------------------\n");
  printf("好きなアルゴリズムで圧縮してください\n");
  printf("[1] ランレングス符号化A 単純にデータをラン長+データ値で圧縮します。\n");
  printf("[2] ランレングス符号化B 二つ以上続くデータ値のみラン長として符号化します。\n");
  printf("[3] ランレングス符号化C ラン長の1bit目でランレングス符号の始まりを示すように符号化します。\n");
  printf("[4] ワイル符号化 符号の先頭に0で終わる1の並びによって、ラン長を示すビット数を決定します。\n");
  printf("[5] PackBits ラン長を示す数が正のとき、同じデータ値が並び、負のとき、違いデータ値が並んでるとして、圧縮します。\n");
  printf("[6] PackBits 同じデータ値が連続して並ぶとき、違うデータ値が連続して並ぶときは間にラン長0があるとして圧縮します。\n");
}

static void selectCompFunc(comp_t *paramsPtr){
  switch(paramsPtr->event){
    case RUNLENGTH_STD:
      compRunLengthStd(paramsPtr);
      break;
    case RUNLENGTH_1BYTE_SIGN:
      compRunLength1ByteSign(paramsPtr);
      break;
    case RUNLENGTH_HEAD1BIT_SIGN:
      compRunLengthHead1BitSign(paramsPtr);
      break;
    case WYLE_CODING:
      compWyleCoding(paramsPtr);
      break;
    case PACKBITS_STD:
      compPackBitsStd(paramsPtr);
      break;
    case PACKBITS_SWITCH:
      compPackBitsSwitch(paramsPtr);
      break;
    default:
      break;
  }
}

static void selectDecompFunc(comp_t *paramsPtr){
  switch(paramsPtr->event){
    case RUNLENGTH_STD:
      decompRunLengthStd(paramsPtr);
      break;
    case RUNLENGTH_1BYTE_SIGN:
      decompRunLength1ByteSign(paramsPtr);
      break;
    case RUNLENGTH_HEAD1BIT_SIGN:
      decompRunLengthHead1BitSign(paramsPtr);
      break;
    case WYLE_CODING:
      decompWyleCoding(paramsPtr);
      break;
    case PACKBITS_STD:
      decompPackBitsStd(paramsPtr);
      break;
    case PACKBITS_SWITCH:
      decompPackBitsSwitch(paramsPtr);
      break;
    default:
      break;
  }
}

static void releaseCompParams(comp_t *compPtr){
  compPtr->fp_i = NULL; 
  compPtr->fp_o = NULL;
  compPtr->data = NULL;
  compPtr->size = 0;
  compPtr->event = NOT_SELECT_RUNLENGTH;
}
