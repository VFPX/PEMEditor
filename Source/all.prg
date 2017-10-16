 Local laUsed[1], lcClass, lnTablesOpen, lnWorkArea, lnX
      lcClass      = [this]
      lnTablesOpen = Aused(laUsed)
      Store Dbc() To (lcClass + [.dataOldDbc])
      If lnTablesOpen > 0
         Store lnTablesOpen To (lcClass + [.dataOldTablesOpen])
         For lnX = 1 To lnTablesOpen
            lnWorkArea = laUsed(lnX, 2)
            Select (lnWorkArea)
            Store lnWorkArea To (lcClass + [.dataOldWorkArea_] + Transform(lnX))
            Store Dbf() To (lcClass + [.dataOldTable_] + Transform(lnX))
            Store Order() To (lcClass + [.dataOldOrder_] + Transform(lnX))
            Store Descending() To (lcClass + [.dataOldDescending_] + Transform(lnX))
            Store Recno() To (lcClass + [.dataOldRecord_] + Transform(lnX))
            Store Isexclusive() To (lcClass + [.dataOldExclusive_] + Transform(lnX))
            Store Set([Relation]) To (lcClass + [.dataOldRelation_] + Transform(lnX))
            Store Set([Filter]) To (lcClass + [.dataOldFilter_] + Transform(lnX))
            If Alias() # Juststem(Dbf())
               Store Alias() To (lcClass + [.dataOldAlias_] + Transform(lnX))
            Else
               Store [] To (lcClass + [.dataOldAlias_] + Transform(lnX))
            Endif
         Endfor
      Endif