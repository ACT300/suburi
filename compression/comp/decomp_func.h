#ifndef DECOMP_FUNC_H
#define DECOMP_FUNC_H

#include "main.h"

extern int decompRunLengthStd(comp_t*);
extern int decompRunLength1ByteSign(comp_t*);
extern int decompRunLengthHead1BitSign(comp_t*);
extern int decompWyleCoding(comp_t*);
extern int decompPackBitsStd(comp_t*);
extern int decompPackBitsSwitch(comp_t*);

#endif /* DECOMP_FUNC_H */
