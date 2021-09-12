/* 利用ファイルのヘッダ */

/* 自ファイルのヘッダ */
#include "comp_func.h"

/* 圧縮してファイルに出力する */

/* list1-1 方法A */
static int getRunLengthStd(comp_t *compPtr, int x, int *run, int *code){
  int limit;
  int processingPoint;

  limit = 0xff;
  processingPoint = x;
  *run = 1;
  *code = (compPtr->data)[processingPoint];
  processingPoint++;
  while(processingPoint < compPtr->size && *code == (compPtr->data)[processingPoint] && *run < limit){
    processingPoint++;
    (*run)++;
  }
  return processingPoint;
}

void compRunLengthStd(comp_t *compPtr){
  int run;
  int code;
  int processingPoint;

  processingPoint = 0;

  while(processingPoint < compPtr->size){
    processingPoint = getRunLengthStd(compPtr, processingPoint, &run, &code);
    fputc(run, compPtr->fp_o);
    fputc(code, compPtr->fp_o);
  }
}

/* list1-2 方法B */
static int getRunLength1ByteSign(comp_t *compPtr, int x, int *run, int *code){
  int limit;
  int processingPoint;

  limit = 0xff;
  processingPoint = x;
  *run = 1;
  *code = (compPtr->data)[processingPoint];
  processingPoint++;
  while(processingPoint < compPtr->size && *code == (compPtr->data)[processingPoint] && *run < limit){
    processingPoint++;
    (*run)++;
  }
  return processingPoint;
}

void compRunLength1ByteSign(comp_t *compPtr){
  int i;
  int run;
  int code;
  int processingPoint;
  int runlengthFlag;

  runlengthFlag = RUNLENGTH_FLAG;

  processingPoint = 0;
  while(processingPoint < compPtr->size){
    processingPoint = getRunLength1ByteSign(compPtr, processingPoint, &run, &code);
    if(run <= 2){
      if(code != runlengthFlag){
        for(i = 0; i < run; i++){
          fputc(code, compPtr->fp_o);
        }
      }else{
        fputc(runlengthFlag, compPtr->fp_o);
        fputc(run, compPtr->fp_o);
        fputc(code, compPtr->fp_o);
      }
    }else{
      fputc(runlengthFlag, compPtr->fp_o);
      fputc(run, compPtr->fp_o);
      fputc(code, compPtr->fp_o);
    }
  }
}

/* list1-3 方法C */
static int getRunLengthHead1BitSign(comp_t *compPtr, int x, int *run, int *code){
  int limit;
  int processingPoint;

  limit = 0x7f;
  processingPoint = x;
  *run = 1;
  *code = (compPtr->data)[processingPoint];
  processingPoint++;
  while(processingPoint < compPtr->size && *code == (compPtr->data)[processingPoint] && *run < limit){
    processingPoint++;
    (*run)++;
  }
  return processingPoint;
}

void compRunLengthHead1BitSign(comp_t *compPtr){
  int run;
  int code;
  int processingPoint;
  
  processingPoint = 0;
  while(processingPoint < compPtr->size){
    processingPoint = getRunLengthHead1BitSign(compPtr, processingPoint, &run, &code);
    if(run < 2){
      if(code < 0x80){
        fputc(code, compPtr->fp_o);
      }else{
        fputc(0x81, compPtr->fp_o);
        fputc(code, compPtr->fp_o);
      }
    }else{
      fputc(run|0x80, compPtr->fp_o);
      fputc(code, compPtr->fp_o);
    }
  }
}

/* list1-4 方法D ワイル符号化 */

/* long型(4バイト)の値をファイルに出力する シフト演算子によってnumは強制的に4バイトになる */
static void fputLong(long num, FILE *fp){
  fputc((num >> 24) & 0xff, fp);
  fputc((num >> 16) & 0xff, fp);
  fputc((num >> 8)  & 0xff, fp);
  fputc((num)       & 0xff, fp);
}

static int absi(int x){
  if(x < 0){
    return -x;
  }
  return x;
}

static void fputBitInit(wyle_t *wylePtr){
  wylePtr->bdata = 0;
  wylePtr->bits = 0;
}

static void fputBit(int bit, wyle_t *wylePtr, FILE *fp){
  wylePtr->bdata = (wylePtr->bdata << 1) | bit;
  wylePtr->bits++;
  if(wylePtr->bits >= 8){
    fputc(wylePtr->bdata, fp);
    wylePtr->bits = 0;
    wylePtr->bdata = 0;
  }
}

static void flushBit(wyle_t *wylePtr, FILE *fp){
  int i;

  for(i = 0; i < 7; i++){
    fputBit(0, wylePtr, fp);
  }
}

static int putBits(int data, int n, wyle_t *wylePtr, FILE *fp){
  int i;

  for(i = n - 1; i >= 0; i--){
    fputBit((data >> i) & 1, wylePtr, fp);
  }
}

static int getLengthForWyle(int x){
  int len;

  x /= 4;
  len = 0;
  while(x > 0){
    x /= 2;
    len++;
  }
  return len;
}

static void outputWyleCode(int value, wyle_t *wylePtr, FILE *fp){
  int switchflag;
  int abs;
  int length;
  int i;

  switchflag = 1;
  abs = absi(value - 1);
  length = getLengthForWyle(abs);

  for(i = 0; i < length; i++){
    fputBit(switchflag, wylePtr, fp);
  }
  fputBit(1-switchflag, wylePtr, fp);
  length += 2;
  for(i = 0; i < length; i++){
    fputBit(abs & 1, wylePtr, fp);
    abs >>= 1;
  }
}

static int getRunLengthWyleCoding(comp_t *compPtr, int x, int *run, int *code){
  int limit;
  int processingPoint;

  limit = 0xffff;
  processingPoint = x;
  *run = 1;
  *code = (compPtr->data)[processingPoint];
  processingPoint++;
  while(processingPoint < compPtr->size && *code == (compPtr->data)[processingPoint] && *run < limit){
    processingPoint++;
    (*run)++;
  }
  return processingPoint;
}

void compWyleCoding(comp_t *compPtr){
  int run;
  int code;
  int processingPoint;
  wyle_t wyleParams;  

  fputLong(compPtr->size, compPtr->fp_o);
  processingPoint = 0;
  fputBitInit(&wyleParams);
  while(processingPoint < compPtr->size){
    processingPoint = getRunLengthWyleCoding(compPtr, processingPoint, &run, &code);
    outputWyleCode(run, &wyleParams, compPtr->fp_o);
    putBits(code, 8, &wyleParams, compPtr->fp_o);
  }
  flushBit(&wyleParams, compPtr->fp_o);
}

/* list1-5 方法E PackBits */
static void getPackBits(comp_t *compPtr, int point, int *packBitMode, int *length){
  if(point >= compPtr->size - 1){   //1バイトだけ残っているパターン
    *packBitMode = NOT_RUN;
    *length = 1;
    return;
  }else{
    if((compPtr->data)[point] == (compPtr->data)[point + 1]){
      *packBitMode = RUN;
    }else{
      *packBitMode = NOT_RUN;
    }
  }
  if(*packBitMode == RUN){
    *length = 1;
    while(1){
      if(point >= compPtr->size - 1){
        break;
      }
      if((compPtr->data)[point] != (compPtr->data)[point + 1]){
        break;
      }
      (*length)++;
      point++;
    }
  }else{
    *length = 0;  //outputPackBitsCodeのfor文の影響で0スタート
    while(1){
      if(point >= compPtr->size - 1){
        (*length)++;  //NOT_RUNはlengthが0スタートの影響で最後に足す
        break;
      }
      if((compPtr->data)[point] == (compPtr->data)[point + 1]){
        break;
      }
      (*length)++;
      point++;
    }
  }
}

static void outputPackBitsCodeStd(comp_t *compPtr, int mode, int length, int point){
  int i;

  if(mode == RUN){
    fputc(length, compPtr->fp_o);
    fputc((compPtr->data)[point], compPtr->fp_o);
  }else{
    fputc(-length, compPtr->fp_o);
    for(i = 0; i < length; i++){
      fputc((compPtr->data)[point + i], compPtr->fp_o);
    }
  }
}

void compPackBitsStd(comp_t *compPtr){
  int point;
  int mode;
  int length;
  int maxlength;

  maxlength = 127;
  point = 0;

  while(point < compPtr->size){
    getPackBits(compPtr, point, &mode, &length);
    while(length > maxlength){
      outputPackBitsCodeStd(compPtr, mode, maxlength, point);
      point += maxlength;
      length -= maxlength;
    }
    outputPackBitsCodeStd(compPtr, mode, length, point);
    point += length;
  }
}

/* list1-6 方法F PackBist */
static void outputPackBitsCodeSW(comp_t *compPtr, int mode, int length, int point){
  int i;

  fputc(length, compPtr->fp_o);
  if(length > 0){
    if(mode == RUN){
      fputc((compPtr->data)[point], compPtr->fp_o);
    }else{
      for(i = 0; i < length; i++){
        fputc((compPtr->data)[point + i], compPtr->fp_o);
      }
    }
  }
}

void compPackBitsSwitch(comp_t *compPtr){
  int point;
  int mode;
  int length;
  int previousMode;
  int maxlength;

  previousMode = NOT_RUN;
  maxlength = 255;

  point = 0;
  while(point < compPtr->size){
    getPackBits(compPtr, point, &mode, &length);
    if(previousMode == mode){
      outputPackBitsCodeSW(compPtr, mode, 0, point);
    }
    while(length > maxlength){
      outputPackBitsCodeSW(compPtr, mode, maxlength, point);
      outputPackBitsCodeSW(compPtr, 1 - mode, 0, point);
      point += maxlength;
      length -= maxlength;
    }
    if(length > 0){
      outputPackBitsCodeSW(compPtr, mode, length, point);
    }
    point += length;
    previousMode = mode;
  }
}
