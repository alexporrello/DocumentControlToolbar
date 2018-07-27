using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentControlToolbar {
    using System;

    public class CustomExceptions : Exception {
        public CustomExceptions() { }
        public CustomExceptions(string message) : base(message) { }
        public CustomExceptions(string message, Exception inner) : base(message, inner) { }
    }

    public class CouldNotDownloadFileException : Exception {
        public CouldNotDownloadFileException() { }
        public CouldNotDownloadFileException(string message) : base(message) { }
        public CouldNotDownloadFileException(string message, Exception inner) : base(message, inner) { }
    }
}
