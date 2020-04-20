# pcr-plugin
Excel Add-In that simulates PCR products.

### Prompt

Make a Visual basic Excel plugin version of PCRSimulator, GoldenGateSimulator, or GibsonSimulator, etc.  PCR simulator is invoked by saying =PCR(A1,A2,A3) where cells A1, A2, and A3 contain the sequence of primer1, primer2, and template.  It should throw an Exception if the sequences are not DNA [ATCGatcg] or the degeneracy code letters like N and K (etc). It should assume that the template is double-stranded/circular and be capable of doing IPCRs. It should find the best annealing sites in the template for the 3' ends of the oligos and calculate the final product sequence.  It can throw an exception if the annealing ~20bp of an oligo contains degeneracy codes (it is a no-no to put degenerate bases into the 3' end of the oligo for this experiment).  A similar algorithm is available in Java as a reference, or to be wrapped if that is possible.  See https://www.reddit.com/r/java/comments/8kdj9v/write_excel_addins_in_java/.
