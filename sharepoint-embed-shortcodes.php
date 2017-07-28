<?php
/*
Plugin Name: SharePoint Document Embed
Plugin URI: https://github.com/BellevueCollege/
Description: Embed publicly shared documents from SharePoint and OneDrive for Business in WordPress
Author: Bellevue College Integration Team
Version: 0.0.0.1
Author URI: http://www.bellevuecollege.edu
*/

defined( 'ABSPATH' ) || exit;

// Shortcode Setup
function spembed_excel_shortcode( $sc_config ) {

	// Shortcode Attributes
	$sc_config = shortcode_atts( array(
		'source'      => false,
		'width'       => '100%',
		'height'      => '500px',
		'bifeatures'  => true,
		'interactive' => false,
		'download'    => true,

	), $sc_config, 'spembed_excel_shortcode' );

	// Allowed hostnames
	$allowed_hosts = array(
		'bellevuec.sharepoint.com',
		'bellevuec-my.sharepoint.com',
	);

	// Parse URL to validate it is from one of the allowed hosts
	$parsed_source_url = wp_parse_url( $sc_config['source'] );
	if ( in_array( $parsed_source_url['host'], $allowed_hosts, true ) ) {

		// Build Request URL
		$request_url = $sc_config['source'] .
			'&action=embedview&wdbipreview=' . $sc_config['bifeatures'] .
			'&wdAllowInteractivity=' . $sc_config['interactive'] .
			'&wdAllowInteractivity=' . $sc_config['interactive'] .
			'&wdDownloadButton=' . $sc_config['download'];

		// Attempt to hide the download link if perameter set
		if ( $sc_config['download'] ) {
			$download_link = '<small><a href="' . $sc_config['source'] . '" target="_blank">Open this spreadsheet in a new window <span class="glyphicon glyphicon-new-window" aria-hidden="true"></span></a></small>';
		} else {
			$download_link = '<a class="sr-only" href="' . $sc_config['source'] . '" target="_blank">Open this spreadsheet in a new window</a>';
		}

		// Return iFrame with embedded document
		return '<iframe
			aria-hidden="true" 
			name="Embedded Excel Document" 
			width="' . $sc_config['width'] . '"
			height="' . $sc_config['height'] . '"
			scrolling="no" 
			src="' . esc_url( $request_url ) . '"></iframe>
			' . $download_link;
	} else {

		// Return error message
		return '<p class="alert alert-danger"><strong>This document can not be displayed</strong> It doesn\'t look like this document is on Bellevue College\'s SharePoint or OneDrive for Business, and it can not be displayed.</p>';
	}
}

add_shortcode( 'sp-excel-embed', 'spembed_excel_shortcode' );


// Add ShortCake UI Support
function register_fusion_pullquote_ui() {
	shortcode_ui_register_for_shortcode(
		'sp-excel-embed',
		array(
			'label' => 'Excel Document',
			'listItemImage' => 'dashicons-grid-view',
			'attrs' => array(
				array(
					'label' => 'Public Sharing URL',
					'description' => 'Public Sharing URL for an Excel document in Bellevue College\'s MyBC SharePoint or OneDrive for Business',
					'attr'  => 'source',
					'type'  => 'url',
					'default' => 'URL',
					'meta' => array(
						'style' => 'width: 100%',
					),
				),
				array(
					'label' => 'Width',
					'attr' => 'width',
					'type' => 'text',
					'meta' => array(
						'placeholder' => 'Recommended: 100%',
					),
				),
				array(
					'label' => 'Height',
					'attr' => 'hight',
					'type' => 'text',
					'meta' => array(
						'placeholder' => 'Recommended: 500px',
					),
				),
				array(
					'label' => 'Business Intelligence Features',
					'attr' => 'bifeatures',
					'type' => 'checkbox',
				),
				array(
					'label' => 'Interactive Features',
					'attr' => 'interactive',
					'type' => 'checkbox',
				),
				array(
					'label' => 'Download Button',
					'attr' => 'download',
					'type' => 'checkbox',
				),
			),
		)
	);
}

add_action( 'init', 'register_fusion_pullquote_ui' );
