#if !defined(DRAGDROPTIMEOFDAY_H)
#define DRAGDROPTIMEOFDAY_H

#include <graphics.hpp>

// TTimeOfDay is the structure which is transferred from the drop source to
// the drop target.
struct TTimeOfDay {
  unsigned short hours,minutes,seconds,milliseconds;
  TColor color;
};

// Name of our custom clipboard format.
#define sTimeOfDayName "TimeOfDay"

#endif
