do ->
  # The initialize function must be run each time a new page is loaded
  run = ->
    ###*
    # Insert your Outlook code here
    ###


  Office.initialize = (reason) ->
    $(document).ready ->
      $('#run').click run
