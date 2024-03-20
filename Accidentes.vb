Sub main
        Stop
        If Not NewValue Is Nothing Then
		'si selecciona el campo otra vez, limpia
			'causa raiz
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = false
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = false
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = false
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = false
                    owner.FechaPrevista     = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = false
                    owner.FechaConclusion   = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = false
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = false
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = false
                'liquidacion
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = false
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = false
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = false
                'siniestro
                    owner.NROSINIESTRO        = ""
                    Owner.Attributes("NROSINIESTRO").ReadOnly = false
                    Set owner.ART                 = Nothing
                    Owner.Attributes("ART").ReadOnly = false
                    owner.AUTODENUNCIA        = False
                    Owner.Attributes("AUTODENUNCIA").ReadOnly = false
                    owner.DIASDEBAJATOTALES   = 0
                    Owner.Attributes("DIASDEBAJATOTALES").ReadOnly = false
                    owner.DIASBAJALABORAL     = 0
                    Owner.Attributes("DIASBAJALABORAL").ReadOnly = false
                    owner.LESION              = ""
                    Owner.Attributes("LESION").ReadOnly = false
                    Set owner.UNIDAD              = Nothing
                    Owner.Attributes("UNIDAD").ReadOnly = false
                    Set owner.SECTORACC           = Nothing
                    Owner.Attributes("SECTORACC").ReadOnly = false
                    Set owner.FORMAACCIDENTE      = Nothing
                    Owner.Attributes("FORMAACCIDENTE").ReadOnly = false
                    Set owner.ZONACUERPO          = Nothing
                    Owner.Attributes("ZONACUERPO").ReadOnly = false
                    Set owner.TIPOACCIDENTE       = Nothing
                    Owner.Attributes("TIPOACCIDENTE").ReadOnly = false
                    owner.GESTIONRECHAZO      = False
                    Owner.Attributes("GESTIONRECHAZO").ReadOnly = false
                    owner.GESTIONEFECTIVADERECHAZO      = false
                    Owner.Attributes("GESTIONEFECTIVADERECHAZO").ReadOnly = false
				'legales
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = false
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = false
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = false
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = false
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = false
					Set owner.TipoAccidente     = Nothing
                    Owner.Attributes("TipoAccidente").ReadOnly = false

			Select Case NewValue.value.id
                Case "{E85B9A48-B7A8-442D-9101-3E0770FA5F0A}", "{E46C5B65-2E9D-41B0-8245-51B8BFDB08AC}"     ' DE TRABAJO / ' DE TRABAJO (RECHAZADO)
                    Set owner.TipoAccidente     = Nothing
                    Owner.Attributes("TipoAccidente").ReadOnly = True
                    owner.NroJuicio         = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio= ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART       = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado  = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True
                    owner.FechaAccidente    = #30/12/1899#
                    Owner.Attributes("FechaAccidente").ReadOnly = True

                    owner.CostoEmpresa      = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero     = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal        = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
                Case    "{881B193C-F10B-4914-913A-86B319927E71}", "{0F996176-A3EA-456B-9381-DEDDC8C16EAE}"  'IN ITINERE/    IN ITINERE (RECHAZADO)
                    Set owner.Unidad            = Nothing
                    Owner.Attributes("Unidad").ReadOnly = True
                    Set owner.SectorAcc         = Nothing
                    Owner.Attributes("SectorAcc").ReadOnly = True
                    owner.DiasBajaLaboral       = ""
                    Owner.Attributes("DiasBajaLaboral").ReadOnly = True
                    Set owner.FormaAccidente    = Nothing
                    Owner.Attributes("FormaAccidente").ReadOnly = True
                    Set owner.ZonaCuerpo        = Nothing
                    Owner.Attributes("ZonaCuerpo").ReadOnly = True
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista         = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion       = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
					'legales
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True

                Case    "{C338FD8E-6612-4C29-A0EE-920EA9500EFF}", "{6287B24C-1E9C-45F2-A5BD-F4A9BCF63641}"  'ENFERMEDAD PROFESIONAL/    ENFERMEDAD PROFESIONAL (RECHAZADA)
                    Set owner.Unidad            = Nothing
                    Owner.Attributes("Unidad").ReadOnly = True
                    Set owner.SectorAcc         = Nothing
                    Owner.Attributes("SectorAcc").ReadOnly = True
                    Set owner.FormaAccidente    = Nothing
                    Owner.Attributes("FormaAccidente").ReadOnly = True
                    Set owner.ZonaCuerpo        = Nothing
                    Owner.Attributes("ZonaCuerpo").ReadOnly = True

                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista         = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion       = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True

                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True

                    owner.NroJuicio         = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio= ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART       = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado  = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True

                Case    "{1AC31F09-B976-4B6D-B9AA-57F99BA4BBE3}", "{84E9B23B-BAFB-4B37-B832-30ECD55D9015}", "{5827271D-6B43-471F-8E79-3EDEB6A26E7B}","{24676CE0-3E9D-4C4F-AEAE-CB58C03D0B5C}", "{91E8FEC2-F969-407F-B06A-6B91E0DDAACA}" , "{E37E4EBB-57E9-4E3A-9DEB-E8D9265B114C}"   'REINGRESOS / REINGRESOS RECHAZADOS
                    Set owner.TipoAccidente     = Nothing
                    Owner.Attributes("TipoAccidente").ReadOnly = True
                    Set owner.Unidad            = Nothing
                    Owner.Attributes("Unidad").ReadOnly = True
                    Set owner.SectorAcc         = Nothing
                    Owner.Attributes("SectorAcc").ReadOnly = True
                    Set owner.FormaAccidente    = Nothing
                    Owner.Attributes("FormaAccidente").ReadOnly = True
                    Set owner.ZonaCuerpo        = Nothing
                    Owner.Attributes("ZonaCuerpo").ReadOnly = True
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista         = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion       = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True
					'legales
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True
				Case "{9451AC1F-6797-475F-85FE-A885E02B5E6B}"  'Accidente no denunciado
                'causa raiz
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista     = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion   = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True
                'liquidacion
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
                'siniestro
                    owner.NROSINIESTRO        = ""
                    Owner.Attributes("NROSINIESTRO").ReadOnly = True
                    Set owner.ART                 = Nothing
                    Owner.Attributes("ART").ReadOnly = True
                    owner.AUTODENUNCIA        = False
                    Owner.Attributes("AUTODENUNCIA").ReadOnly = True
                    owner.DIASDEBAJATOTALES   = 0
                    Owner.Attributes("DIASDEBAJATOTALES").ReadOnly = True
                    owner.DIASBAJALABORAL     = 0
                    Owner.Attributes("DIASBAJALABORAL").ReadOnly = True
                    owner.LESION              = ""
                    Owner.Attributes("LESION").ReadOnly = True
                    Set owner.UNIDAD              = Nothing
                    Owner.Attributes("UNIDAD").ReadOnly = True
                    Set owner.SECTORACC           = Nothing
                    Owner.Attributes("SECTORACC").ReadOnly = True
                    Set owner.FORMAACCIDENTE      = Nothing
                    Owner.Attributes("FORMAACCIDENTE").ReadOnly = True
                    Set owner.ZONACUERPO          = Nothing
                    Owner.Attributes("ZONACUERPO").ReadOnly = True
                    Set owner.TIPOACCIDENTE       = Nothing
                    Owner.Attributes("TIPOACCIDENTE").ReadOnly = True
                    owner.GESTIONRECHAZO      = False
                    Owner.Attributes("GESTIONRECHAZO").ReadOnly = True
                    owner.GESTIONEFECTIVADERECHAZO      = false
                    Owner.Attributes("GESTIONEFECTIVADERECHAZO").ReadOnly = True
				'legales
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True
	
                Case "{D0D6B3FA-64D4-4581-BBC0-A9B50062EF91}" 'juicio ART
                'causa raiz
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista     = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion   = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True
                'liquidacion
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
                'siniestro
                    owner.NROSINIESTRO      = ""
                    Owner.Attributes("NROSINIESTRO").ReadOnly = True
                    owner.AUTODENUNCIA        = False
                    Owner.Attributes("AUTODENUNCIA").ReadOnly = True
                    owner.DIASDEBAJATOTALES   = 0
                    Owner.Attributes("DIASDEBAJATOTALES").ReadOnly = True
                    owner.DIASBAJALABORAL     = 0
                    Owner.Attributes("DIASBAJALABORAL").ReadOnly = True
                    owner.LESION              = ""
                    Owner.Attributes("LESION").ReadOnly = True
                    Set owner.UNIDAD              = Nothing
                    Owner.Attributes("UNIDAD").ReadOnly = True
                    Set owner.SECTORACC           = Nothing
                    Owner.Attributes("SECTORACC").ReadOnly = True
                    Set owner.FORMAACCIDENTE      = Nothing
                    Owner.Attributes("FORMAACCIDENTE").ReadOnly = True
                    Set owner.ZONACUERPO          = Nothing
                    Owner.Attributes("ZONACUERPO").ReadOnly = True
                    Set owner.TIPOACCIDENTE       = Nothing
                    Owner.Attributes("TIPOACCIDENTE").ReadOnly = True
                    owner.GESTIONRECHAZO      = False
                    Owner.Attributes("GESTIONRECHAZO").ReadOnly = True
                    owner.GESTIONEFECTIVADERECHAZO      = false
                    Owner.Attributes("GESTIONEFECTIVADERECHAZO").ReadOnly = True
                Case "{2CC2C99B-7BE3-4292-ACB2-643BFE038100}" , "{ADFB92A5-1460-48C4-B151-201AFE0F1E48}" , "{F2D4A583-3710-4775-B8F9-57EBFDD3514B}" 'post desvinculacion
                'legales
                    owner.NroJuicio             = ""
                    Owner.Attributes("NroJuicio").ReadOnly = True
                    Set owner.EstadoDeJucio    = Nothing
                    Owner.Attributes("EstadoDeJucio").ReadOnly = True
                    owner.JuridiccionJuicio    = ""
                    Owner.Attributes("JuridiccionJuicio").ReadOnly = True
                    owner.AnalistaART           = ""
                    Owner.Attributes("AnalistaART").ReadOnly = True
                    owner.LetradoDesignado      = ""
                    Owner.Attributes("LetradoDesignado").ReadOnly = True
				'causa raiz
                    Set owner.CausaRaiz         = Nothing
                    Owner.Attributes("CausaRaiz").ReadOnly = True
                    Set owner.SubCausa          = Nothing
                    Owner.Attributes("SubCausa").ReadOnly = True
                    Set owner.Accion            = Nothing
                    Owner.Attributes("Accion").ReadOnly = True
                    owner.Responsable       = ""
                    Owner.Attributes("Responsable").ReadOnly = True
                    owner.FechaPrevista     = #30/12/1899#
                    Owner.Attributes("FechaPrevista").ReadOnly = True
                    owner.FechaConclusion   = #30/12/1899#
                    Owner.Attributes("FechaConclusion").ReadOnly = True
                    Set owner.Status            = Nothing
                    Owner.Attributes("Status").ReadOnly = True
                    Set owner.ControlExtension  = Nothing
                    Owner.Attributes("ControlExtension").ReadOnly = True
                'liquidacion
                    owner.CostoEmpresa          = 0.0
                    Owner.Attributes("CostoEmpresa").ReadOnly = True
                    owner.CostoRecupero         = 0.0
                    Owner.Attributes("CostoRecupero").ReadOnly = True
                    owner.CostoTotal            = 0.0
                    Owner.Attributes("CostoTotal").ReadOnly = True
                'siniestro
                    owner.AUTODENUNCIA        = False
                    Owner.Attributes("AUTODENUNCIA").ReadOnly = True
                    owner.DIASDEBAJATOTALES   = 0
                    Owner.Attributes("DIASDEBAJATOTALES").ReadOnly = True
                    owner.DIASBAJALABORAL     = 0
                    Owner.Attributes("DIASBAJALABORAL").ReadOnly = True
                    owner.LESION              = ""
                    Owner.Attributes("LESION").ReadOnly = True
                    Set owner.UNIDAD              = Nothing
                    Owner.Attributes("UNIDAD").ReadOnly = True
                    Set owner.SECTORACC           = Nothing
                    Owner.Attributes("SECTORACC").ReadOnly = True
                    Set owner.FORMAACCIDENTE      = Nothing
                    Owner.Attributes("FORMAACCIDENTE").ReadOnly = True
                    Set owner.ZONACUERPO          = Nothing
                    Owner.Attributes("ZONACUERPO").ReadOnly = True
                    Set owner.TIPOACCIDENTE       = Nothing
                    Owner.Attributes("TIPOACCIDENTE").ReadOnly = True
            End Select
        End If
end sub




