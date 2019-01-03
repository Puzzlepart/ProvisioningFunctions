﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cumulus.Monads
{
    public class Result<T>
    {
        public Result(T value)
        {
            Value = value;
        }
        public T Value { get; set; }
    }
}
