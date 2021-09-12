#ifndef MAIN_H
#define MAIN_H

#include <stdio.h>
#include <stdlib.h>

#define NOT_RUN 1   //同じ値が連続していない部分を示す
#define RUN 0       //同じ値が連続している部分を示す
#define BUFFER_SIZE 1024*1024

#define RUNLENGTH_FLAG 0xff

typedef unsigned char byte;

typedef struct{
  FILE *fp_o;
  FILE *fp_i;
  byte *data;
  int size;
  char event;
}comp_t;

typedef struct{
  int bdata;
  int bits;
}wyle_t;

typedef enum{
  NOT_SELECT_RUNLENGTH = 0x00,
  RUNLENGTH_STD,
  RUNLENGTH_1BYTE_SIGN,
  RUNLENGTH_HEAD1BIT_SIGN,
  WYLE_CODING,
  PACKBITS_STD,
  PACKBITS_SWITCH,
  RUNLENGTH_NUM,
}runlength_t;

#endif /* MAIN_H */
