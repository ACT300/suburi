#ifndef COMP_FUNC_H
#define COMP_FUNC_H

#include "main.h"

extern void compRunLengthStd(comp_t*);
extern void compRunLength1ByteSign(comp_t*);
extern void compRunLnegthHead1BitSign(comp_t*);
extern void compWyleCoding(comp_t*);
extern void compPackBitsStd(comp_t*);
extern void compPackBitsSwitch(comp_t*);

#endif /* COMP_FUNC_H */
