/* 利用ファイルのヘッダ */

/* 自ファイルのヘッダ */
#include "decomp_func.h"

/* 圧縮したファイルの伸張をしてからファイルに出力する */

/* list1-7 方法A */
int decompRunLengthStd(comp_t *paramsPtr){
  int i;
  int code;
  int length;

  while(1){
    length = fgetc(paramsPtr->fp_i);
    if(length == EOF){
      break;
    }
    code = fgetc(paramsPtr->fp_i);
    if(code == EOF){
      return 0; //エラー
    }
    for(i = 0; i < length; i++){
      fputc(code, paramsPtr->fp_o);
    }
  }
  return 1;
}

/* list1-8 方法B */
int decompRunLength1ByteSign(comp_t *paramsPtr){
  int i;
  int length;
  int code;
  int value;
  int runlengthFlag;

  runlengthFlag = RUNLENGTH_FLAG;

  while(1){
    value = fgetc(paramsPtr->fp_i);
    if(value == EOF){
      break;
    }

    if(value == runlengthFlag){
      length = fgetc(paramsPtr->fp_i);
      if(length == EOF){
        return 0; //エラー
      }
      code = fgetc(paramsPtr->fp_i);
      if(code == EOF){
        return 0; //エラー
      }
      for(i = 0; i < length; i++){
        fputc(code, paramsPtr->fp_o);
      }
    }else{
      fputc(value, paramsPtr->fp_o);
    }
  }
  return 1;
}

/* list1-9 方法C */
int decompRunLengthHead1BitSign(comp_t *paramsPtr){
  int i;
  int length;
  int code;
  int value;

  while(1){
    value = fgetc(paramsPtr->fp_i);
    if(value == EOF){
      break;
    }
    if(value >= 0x80){
      length = value & 0x7f;
      code = fgetc(paramsPtr->fp_i);
      if(code == EOF){
        return 0; //エラー
      }
      for(i = 0; i < length; i++){
        fputc(code, paramsPtr->fp_o);
      }
    }else{
      fputc(value, paramsPtr->fp_o);
    }
  }
  return 1;
}

/* list1-10 方法D */
static int readData(unsigned char *data, int length, FILE *fp){
  return fread(data, length, sizeof(char), fp);
}

static int fgetLong(FILE *fp){
  unsigned char buf[16];

  readData(buf, 4, fp);
  return (buf[0] << 24) | (buf[1] << 16) | (buf[2] << 8) | buf[3];
}

static int fgetBitInit(wyle_t *wylePtr){
  wylePtr->bits = 0;
  wylePtr->bdata = 0;
}

static int fgetBit(FILE *fp, wyle_t *wylePtr){
  unsigned char bbuf;
  unsigned char buf[16];
  int val;

  if(wylePtr->bits == 0){
    readData(&bbuf, 1, fp);
    wylePtr->bdata = bbuf;
  }
  val = (wylePtr->bdata >> 7) & 1;
  wylePtr->bdata = (wylePtr->bdata << 1) & 0xff;
  wylePtr->bits++;
  if(wylePtr->bits >= 8){
    wylePtr->bits = 0;
    wylePtr->bdata = 0;
  }

  return val;
}

static int freadWyleCode(FILE *fp, wyle_t *wylePtr){
  int sbit;
  int bit;
  int i;
  int cnt;
  int val;

  cnt = 0;
  while(fgetBit(fp, wylePtr) == 1){
    cnt++;
  }
  cnt += 2;
  val = 0;
  for(i = 0; i < cnt; i++){
    bit = fgetBit(fp, wylePtr) << i;
    val |= bit;
  }

  return val + 1;
}

static int fgetBits(int n, FILE *fp, wyle_t *wylePtr){
  int i;
  int answer;

  answer = 0;

  for(i = 0; i < n; i++){
    answer *= 2;
    answer += fgetBit(fp, wylePtr);
  }

  return answer;
}

int decompWyleCoding(comp_t *paramsPtr){
  int i;
  int point;
  int datasize;
  int code;
  int length;
  wyle_t wyleParams;

  datasize = fgetLong(paramsPtr->fp_i);
  point = 0;
  fgetBitInit(&wyleParams);
  while(point < datasize){
    length = freadWyleCode(paramsPtr->fp_i, &wyleParams);

    code = fgetBits(8, paramsPtr->fp_i, &wyleParams);
    for(i = 0; i < length; i++){
      fputc(code, paramsPtr->fp_o);
    }
    point += length;
  }
  return 1;
}

/* list1-11 方法E */
int decompPackBitsStd(comp_t *paramsPtr){
  int i;
  int code;
  int length;

  while((length = fgetc(paramsPtr->fp_i)) != EOF){
    if(length > 128){ 
      /*int lengthが129(1000 0001)の場合 maxlengthは127だがNOT RUNは負数になっているので 実態は127(0111 1111)*/
      /* プラスの値をマイナスの値にする元々の流れ */
      /* 1. 127 0111 1111 intの世界では129 */
      /* 2.     1000 0000 */
      /* 3.     1000 0001 で -127 でlengthとなる*/
      /* 256 - length(129) = 127  */
      length = 256 - length;
      for(i = 0; i < length; i++){
        code = fgetc(paramsPtr->fp_i);
        if(code == EOF){
          return 0; //エラー
        }
        fputc(code, paramsPtr->fp_o);
      }
    }else{
      /* lengthが127(0111 1111)の場合 */
      code = fgetc(paramsPtr->fp_i);
      if(code == EOF){
        return 0;
      }
      for(i = 0; i < length; i++){
        fputc(code, paramsPtr->fp_o);
      }
    }
  }
  return 1;
}

/* list1-12 方法F */
int decompPackBitsSwitch(comp_t *paramsPtr){
  int i;
  int code;
  int packBitsMode;
  int length;

  packBitsMode = RUN;

  while((length = fgetc(paramsPtr->fp_i)) != EOF){
    if(length > 0){
      if(packBitsMode == NOT_RUN){
        for(i = 0; i < length; i++){
          code = fgetc(paramsPtr->fp_i);
          if(code == EOF){
            return 0; //エラー
          }
          fputc(code, paramsPtr->fp_o);
        }
      }else{
        code = fgetc(paramsPtr->fp_i);
        if(code == EOF){
          return 0; //エラー
        }
        for(i = 0; i < length; i++){
          fputc(code, paramsPtr->fp_o);
        }
      }
    }
    /* 原則 NOT_RUN,RUN を交互に繰り返すが、RUNが続くときは0を間に挟むことでpackBitsModeが0を挟む前と同じになる */
    packBitsMode = 1 - packBitsMode;
  }
  return 1;
}
