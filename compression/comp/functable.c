typedef void (comp_func)(comp_t *);

comp_func *const compFunc[] = {
  compRunLengthStd,
  compRunLength1ByteSign,
  compRunLengthHead1BitSign,
  compWyleCoding,
  compPackBitsStd,
  compPackBitsSwitch,
};

typedef int (decomp_func)(comp_t *);

decomp_func *const decompFunc[] = {
  decompRunLengthStd,
  decompRunLength1ByteSign,
  decompRunLengthHead1BitSign,
  decompWyleCoding,
  decompPackBitsStd,
  decompPackBitsSwitch,
};

enum {
  F_RUNLENGTH_STD = 0x00,
  F_RUNLENGTH_1BYTE_SIGN,
  F_RUNLENGTH_HEAD1BIT_SIGN,
  F_WYLE_CODING,
  F_PACKBITS_STD,
  F_PACKBITS_SWITCH,
  F_RUNLENGTH_NUM,
};

int func_num = F_RUNLENGTH_STD
int status;
status = (decompFunc[func_num])(paramsPtr);
