#########################################################################################
# WorkflowCommander, copyrighted by Joel Wiesmann, 2016
# 
# Warm welcome to my code, whatever wisdom you try to find here.
#
# This file is part of WorkflowCommander. See the manual for details.
#
# THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, 
# INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
# A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT,
# INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
# LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
# OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY
# WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#########################################################################################
$VerbosePreference = 'Continue'

#########################################################################################
# Enums
#########################################################################################

# PSMs seem to have no proper way to export enums that have been introduced with PS5.

Enum WFCSearchObjTypes {
  JOBS
  JOBP
  CALE
  CALL
  CITC
  CLNT
  CODE
  CONN
  CPIT
  DASH
  DOCU
  EVNT
  FILTER
  FOLD
  HOST
  HOSTG
  HSTA
  JOBF
  JOBG
  JOBI
  JOBQ
  JSCH
  LOGIN
  PERIOD
  PRPT
  QUEUE
  SCRI
  SERV
  STORE
  SYNC
  TZ
  USER
  USERG
  VARA
  XSL
  Executeable
}
