using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentControlToolbar {
    using System;

    public class AcronymTableNotFoundException : Exception {
        public AcronymTableNotFoundException() {

        }

        public AcronymTableNotFoundException(string message) : base(message) {
        }

        public AcronymTableNotFoundException(string message, Exception inner) : base(message, inner) {
        }
    }
}
